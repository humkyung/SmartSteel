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

// MainFrm.cpp : implementation of the CMainFrame class
//

#include "stdafx.h"
#include <assert.h>
#include "SmartSteel.h"
#include "AppDocData.h"
#include "Command/AppCommandManager.h"

#include "MainFrm.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

// CMainFrame

IMPLEMENT_DYNCREATE(CMainFrame, CFrameWndEx)

BEGIN_MESSAGE_MAP(CMainFrame, CFrameWndEx)
	ON_WM_CREATE()
	ON_COMMAND(ID_FILE_PRINT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_DIRECT, &CMainFrame::OnFilePrint)
	ON_COMMAND(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnFilePrintPreview)
	ON_MESSAGE(DISPLAY_MESSAGE , &CMainFrame::OnDisplayMessage)
	ON_MESSAGE(DISPLAY_STATUSBAR , &CMainFrame::OnDisplayStatusBar)
	ON_UPDATE_COMMAND_UI(ID_FILE_PRINT_PREVIEW, &CMainFrame::OnUpdateFilePrintPreview)
	ON_UPDATE_COMMAND_UI(ID_EDIT_UNDO, &CMainFrame::OnUpdateEditUndo)
	ON_COMMAND(ID_EDIT_UNDO, &CMainFrame::OnEditUndo)
END_MESSAGE_MAP()

// CMainFrame construction/destruction

static CMainFrame* __instance__ = NULL;

CMainFrame::CMainFrame()
{
	__instance__ = this;
}

CMainFrame::~CMainFrame()
{
}

CMainFrame* CMainFrame::GetInstance()
{
	return __instance__;
}

int CMainFrame::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CFrameWndEx::OnCreate(lpCreateStruct) == -1)
		return -1;

	BOOL bNameValid;

	// set the visual manager used to draw all user interface elements
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));

	// set the visual style to be used the by the visual manager
	CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_ObsidianBlack);

	m_wndRibbonBar.Create(this);
	InitializeRibbon();

	if (!m_wndStatusBar.Create(this))
	{
		TRACE0("Failed to create status bar\n");
		return -1;      // fail to create
	}

	CString strTitlePane1 , strTitlePane2 , strTitlePane3;
	bNameValid = strTitlePane1.LoadString(IDS_STATUS_PANE1);
	ASSERT(bNameValid);
	bNameValid = strTitlePane2.LoadString(IDS_STATUS_PANE2);
	ASSERT(bNameValid);
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		strTitlePane2.Format(_T("Ver : %s") , docData.GetFileVersion());
	}
	bNameValid = strTitlePane3.LoadString(IDS_STATUS_PANE3);
	ASSERT(bNameValid);
	m_wndStatusBar.AddElement(new CMFCRibbonStatusBarPane(IDS_STATUS_PANE1, strTitlePane1, TRUE), strTitlePane1);
	m_wndStatusBar.AddElement(new CMFCRibbonStatusBarPane(IDS_STATUS_PANE2, strTitlePane2, TRUE), strTitlePane2);
	m_wndStatusBar.AddExtendedElement(new CMFCRibbonStatusBarPane(IDS_STATUS_PANE3, strTitlePane3, TRUE), strTitlePane3);
	
	/// enable Visual Studio 2005 style docking window behavior
	CDockingManager::SetDockingMode(DT_SMART);
	/// enable Visual Studio 2005 style docking window auto-hide behavior
	EnableAutoHidePanes(CBRS_ALIGN_ANY);

	// Load menu item image (not placed on any standard toolbars):
	CMFCToolBar::AddToolBarForImageCollection(IDR_MENU_IMAGES, theApp.m_bHiColorIcons ? IDB_MENU_IMAGES_24 : 0);

	// create docking windows
	if (!CreateDockingWindows())
	{
		TRACE0("Failed to create docking windows\n");
		return -1;
	}

	m_wndFileView.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndFileView);
	CDockablePane* pTabbedBar = NULL;
	m_wndOutput.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndOutput);
	m_wndProperties.EnableDocking(CBRS_ALIGN_ANY);
	DockPane(&m_wndProperties);

	return 0;
}

BOOL CMainFrame::PreCreateWindow(CREATESTRUCT& cs)
{
	if( !CFrameWndEx::PreCreateWindow(cs) )
		return FALSE;
	// TODO: Modify the Window class or styles here by modifying
	//  the CREATESTRUCT cs

	return TRUE;
}

/**
	@brief	initialize ribbon
	@author	humkyung
*/
void CMainFrame::InitializeRibbon()
{
	BOOL bNameValid;

	CString strTemp;
	bNameValid = strTemp.LoadString(IDS_RIBBON_FILE);
	ASSERT(bNameValid);

	/// Load panel images:
	m_PanelImages.SetImageSize(CSize(16, 16));
	m_PanelImages.Load(IDB_BUTTONS);

	// Init main button:
	m_MainButton.SetImage(IDB_MAIN);
	m_MainButton.SetText(_T("\nf"));
	m_MainButton.SetToolTipText(strTemp);

	m_wndRibbonBar.SetApplicationButton(&m_MainButton, CSize (45, 45));
	CMFCRibbonMainPanel* pMainPanel = m_wndRibbonBar.AddMainCategory(strTemp, IDB_FILESMALL, IDB_FILELARGE);

	bNameValid = strTemp.LoadString(IDS_RIBBON_OPEN);
	ASSERT(bNameValid);
	pMainPanel->Add(new CMFCRibbonButton(ID_FILE_OPEN, strTemp, 1, 1));
#ifdef	SMART_STEEL	
	bNameValid = strTemp.LoadString(IDS_RIBBON_SAVE);
	ASSERT(bNameValid);
	pMainPanel->Add(new CMFCRibbonButton(ID_FILE_SAVE, strTemp, 2, 2));
	bNameValid = strTemp.LoadString(IDS_RIBBON_SAVEAS);
	ASSERT(bNameValid);
	pMainPanel->Add(new CMFCRibbonButton(ID_FILE_SAVE_AS, strTemp, 3, 3));
#endif
	pMainPanel->Add(new CMFCRibbonSeparator(TRUE));

	bNameValid = strTemp.LoadString(IDS_RIBBON_CLOSE);
	ASSERT(bNameValid);
	pMainPanel->Add(new CMFCRibbonButton(ID_FILE_CLOSE, strTemp, 9, 9));

	bNameValid = strTemp.LoadString(IDS_RIBBON_RECENT_DOCS);
	ASSERT(bNameValid);
	pMainPanel->AddRecentFilesList(strTemp);

	bNameValid = strTemp.LoadString(IDS_RIBBON_EXIT);
	ASSERT(bNameValid);
	pMainPanel->AddToBottom(new CMFCRibbonMainPanelButton(ID_APP_EXIT, strTemp, 15));

	/// Add "Home" category with "Clipboard" panel:
	bNameValid = strTemp.LoadString(IDS_RIBBON_HOME);
	ASSERT(bNameValid);
	CMFCRibbonCategory* pCategoryHome = m_wndRibbonBar.AddCategory(strTemp, IDB_WRITESMALL, IDB_WRITELARGE);

	/// Create "App" panel:
	CMFCRibbonPanel* pPanelApp = pCategoryHome->AddPanel(_T("App"), m_PanelImages.ExtractIcon(27));

	bNameValid = strTemp.LoadString(IDS_APP_GENERATE);
	ASSERT(bNameValid);
	CMFCRibbonButton* pBtnGenerate = new CMFCRibbonButton(IDS_APP_GENERATE , strTemp, -1 , 2);
	pPanelApp->Add(pBtnGenerate);

	bNameValid = strTemp.LoadString(IDS_APP_OPTIONS);
	ASSERT(bNameValid);
	CMFCRibbonButton* pBtnOptions = new CMFCRibbonButton(IDS_APP_OPTIONS , strTemp, -1 , 3);
	pPanelApp->Add(pBtnOptions);

	bNameValid = strTemp.LoadString(ID_EDIT_UNDO);
	ASSERT(bNameValid);
	CMFCRibbonButton* pBtnUndo = new CMFCRibbonButton(ID_EDIT_UNDO , strTemp, -1 , 9);
	pPanelApp->Add(pBtnUndo);

	/// Create "Export" panel:
	CMFCRibbonPanel* pPanelExport = pCategoryHome->AddPanel(_T("Export"), m_PanelImages.ExtractIcon(27));
	{
#ifdef	SMART_STEEL
		CMFCRibbonButton* pBtnExportBREP = new CMFCRibbonButton(IDS_EXPORT_BREP , _T("BREP"), 19);
		pPanelExport->Add(pBtnExportBREP);

		CMFCRibbonButton* pBtnExportIGES = new CMFCRibbonButton(IDS_EXPORT_IGES , _T("IGES"), 20);
		pPanelExport->Add(pBtnExportIGES);
#endif
		CMFCRibbonButton* pBtnExport = new CMFCRibbonButton(IDS_APP_EXPORT , _T("FWP"), 11);
		pPanelExport->Add(pBtnExport);
	}

	/// Create "OCC View" panel:
	bNameValid = strTemp.LoadString(IDS_RIBBON_OCC_VIEW);
	ASSERT(bNameValid);
	CMFCRibbonPanel* pPanelOCCView = pCategoryHome->AddPanel(strTemp, m_PanelImages.ExtractIcon(27));
	{
#ifdef	SMART_STEEL
		bNameValid = strTemp.LoadString(ID_VIEW_INFO);
		ASSERT(bNameValid);
		CMFCRibbonButton* pBtnInfo = new CMFCRibbonButton(ID_VIEW_INFO , strTemp, -1 , 4);
		pPanelOCCView->Add(pBtnInfo);

		bNameValid = strTemp.LoadString(ID_VIEW_ROTATE);
		ASSERT(bNameValid);
		CMFCRibbonButton* pBtnRotate = new CMFCRibbonButton(ID_VIEW_ROTATE , strTemp, -1 , 5);
		pPanelOCCView->Add(pBtnRotate);

		bNameValid = strTemp.LoadString(ID_ROTATE_ABOUT_AXIS);
		ASSERT(bNameValid);
		CMFCRibbonButton* pBtnRotateAboutAxis = new CMFCRibbonButton(ID_ROTATE_ABOUT_AXIS , strTemp, -1 , 6);
		pPanelOCCView->Add(pBtnRotateAboutAxis);
#else
		bNameValid = strTemp.LoadString(ID_VIEW_ROTATE);
		ASSERT(bNameValid);
		CMFCRibbonButton* pBtnPaste = new CMFCRibbonButton(ID_VIEW_ROTATE , strTemp, -1 , 5);
		pPanelOCCView->Add(pBtnPaste);
#endif

		bNameValid = strTemp.LoadString(ID_VIEW_ZOOM);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_VIEW_ZOOM, strTemp, -1 , 7));

		bNameValid = strTemp.LoadString(ID_VIEW_FIT);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_VIEW_FIT, strTemp, -1 , 8));
			
		bNameValid = strTemp.LoadString(ID_OCC_VIEW_TOP);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_TOP, strTemp, 16));

		bNameValid = strTemp.LoadString(ID_OCC_VIEW_LEFT);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_LEFT, strTemp, 14));
		
		bNameValid = strTemp.LoadString(ID_OCC_VIEW_FRONT);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_FRONT , strTemp, 12));

		bNameValid = strTemp.LoadString(ID_OCC_VIEW_BOTTOM);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_BOTTOM , strTemp, 17));

		bNameValid = strTemp.LoadString(ID_OCC_VIEW_RIGHT);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_RIGHT , strTemp, 15));
		
		bNameValid = strTemp.LoadString(ID_OCC_VIEW_BACK);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_BACK , strTemp, 13));

		bNameValid = strTemp.LoadString(ID_OCC_VIEW_ISO);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(ID_OCC_VIEW_ISO , strTemp, 18));

		pPanelOCCView->Add(new CMFCRibbonSeparator(TRUE));

		bNameValid = strTemp.LoadString(IDS_OCC_VIEW_SHADED);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(IDS_OCC_VIEW_SHADED, strTemp , 21));

		bNameValid = strTemp.LoadString(IDS_OCC_VIEW_WIREFRAME);
		ASSERT(bNameValid);
		pPanelOCCView->Add(new CMFCRibbonButton(IDS_OCC_VIEW_WIREFRAME, strTemp , 22));
	}

	/// Create and add a "View" panel:
	bNameValid = strTemp.LoadString(IDS_RIBBON_VIEW);
	ASSERT(bNameValid);
	CMFCRibbonPanel* pPanelView = pCategoryHome->AddPanel(strTemp, m_PanelImages.ExtractIcon (7));

	bNameValid = strTemp.LoadString(IDS_RIBBON_STATUSBAR);
	ASSERT(bNameValid);
	CMFCRibbonButton* pBtnStatusBar = new CMFCRibbonCheckBox(ID_VIEW_STATUS_BAR, strTemp);
	pPanelView->Add(pBtnStatusBar);

	// Add quick access toolbar commands:
	CList<UINT, UINT> lstQATCmds;

	lstQATCmds.AddTail(ID_FILE_NEW);
	lstQATCmds.AddTail(ID_FILE_OPEN);
	lstQATCmds.AddTail(ID_FILE_SAVE);
	lstQATCmds.AddTail(ID_FILE_PRINT_DIRECT);

	m_wndRibbonBar.SetQuickAccessCommands(lstQATCmds);

	m_wndRibbonBar.AddToTabs(new CMFCRibbonButton(ID_APP_ABOUT, _T("\na"), m_PanelImages.ExtractIcon (0)));
}

BOOL CMainFrame::CreateDockingWindows()
{
	BOOL bNameValid;

	// Create file view
	CString strFileView;
	bNameValid = strFileView.LoadString(IDS_FILE_VIEW);
	ASSERT(bNameValid);
	if (!m_wndFileView.Create(strFileView, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_FILEVIEW, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_LEFT| CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create File View window\n");
		return FALSE; // failed to create
	}

	// Create output window
	CString strOutputWnd;
	bNameValid = strOutputWnd.LoadString(IDS_OUTPUT_WND);
	ASSERT(bNameValid);
	if (!m_wndOutput.Create(strOutputWnd, this, CRect(0, 0, 100, 100), TRUE, ID_VIEW_OUTPUTWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_BOTTOM | CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Output window\n");
		return FALSE; // failed to create
	}

	// Create properties window
	CString strPropertiesWnd;
	bNameValid = strPropertiesWnd.LoadString(IDS_PROPERTIES_WND);
	ASSERT(bNameValid);
	if (!m_wndProperties.Create(strPropertiesWnd, this, CRect(0, 0, 200, 200), TRUE, ID_VIEW_PROPERTIESWND, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | CBRS_RIGHT | CBRS_FLOAT_MULTI))
	{
		TRACE0("Failed to create Properties window\n");
		return FALSE; // failed to create
	}

	SetDockingWindowIcons(theApp.m_bHiColorIcons);
	return TRUE;
}

void CMainFrame::SetDockingWindowIcons(BOOL bHiColorIcons)
{
	HICON hFileViewIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_FILE_VIEW_HC : IDI_FILE_VIEW), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndFileView.SetIcon(hFileViewIcon, FALSE);

	HICON hOutputBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_OUTPUT_WND_HC : IDI_OUTPUT_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndOutput.SetIcon(hOutputBarIcon, FALSE);

	HICON hPropertiesBarIcon = (HICON) ::LoadImage(::AfxGetResourceHandle(), MAKEINTRESOURCE(bHiColorIcons ? IDI_PROPERTIES_WND_HC : IDI_PROPERTIES_WND), IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), 0);
	m_wndProperties.SetIcon(hPropertiesBarIcon, FALSE);

}

// CMainFrame diagnostics

#ifdef _DEBUG
void CMainFrame::AssertValid() const
{
	CFrameWndEx::AssertValid();
}

void CMainFrame::Dump(CDumpContext& dc) const
{
	CFrameWndEx::Dump(dc);
}
#endif //_DEBUG


// CMainFrame message handlers

void CMainFrame::OnFilePrint()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_PRINT);
	}
}

void CMainFrame::OnFilePrintPreview()
{
	if (IsPrintPreview())
	{
		PostMessage(WM_COMMAND, AFX_ID_PREVIEW_CLOSE);  // force Print Preview mode closed
	}
}

/**
	@brief	display StatusBar
	@author	humkyung
	@date	2013.07.27
*/
LRESULT CMainFrame::OnDisplayStatusBar(WPARAM wParam, LPARAM lParam)
{
	m_wndStatusBar.GetElement(0)->SetText((NULL != wParam) ? LPCSTR(wParam) : _T("Ready"));
	m_wndStatusBar.ForceRecalcLayout();
	m_wndStatusBar.GetElement(0)->Redraw();
	
	return 0L;
}

/**
	@brief	display message
	@author	humkyung
	@date	2013.07.09
*/
LRESULT CMainFrame::OnDisplayMessage(WPARAM wParam, LPARAM lParam)
{
	CString str = LPCTSTR(wParam);
	if(str.IsEmpty())
	{
		m_wndOutput.Clear();
	}
	else
	{
		m_wndOutput.DisplayMessage(str , (MessageType::MESSAGE_WARNING == lParam) ? MessageType::MESSAGE_WARNING : MessageType::MESSAGE_INFO);
	}

	return 0L;
}

void CMainFrame::OnUpdateFilePrintPreview(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(IsPrintPreview());
}

/**
	@brief	fill properties of sdnf element attribute
	@author	humkyung
	@date	2013.06.11
*/
int CMainFrame::FillPropertiesOf(CSDNFAttribute* pAttr)
{
	assert(pAttr && "pAttr is NULL");

	if(pAttr)
	{
		return m_wndProperties.FillPropertiesOf(pAttr);
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	fill properties of plate attribute
	@author	humkyung
	@date	2013.08.04
*/
int CMainFrame::FillPropertiesOf(CSteelPlate* pPlate)
{
	assert(pPlate && "pPlate is NULL");

	if(pPlate)
	{
		return m_wndProperties.FillPropertiesOf(pPlate);
	}

	return ERROR_INVALID_PARAMETER;
}

void CMainFrame::OnUpdateEditUndo(CCmdUI *pCmdUI)
{
	Command::CAppCommandManager& man = Command::CAppCommandManager::GetInstance();
	pCmdUI->Enable(man.GetCommandCount() > 0);
}

/**
	@brief	undo the last command
	@author	humkyung
	@date	2013.07.27
*/
void CMainFrame::OnEditUndo()
{
	Command::CAppCommandManager& man = Command::CAppCommandManager::GetInstance();
	man.Undo();
}
