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

// SmartSteel.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "afxwinappex.h"

#include "SmartSteel.h"
#include "MainFrm.h"

#include "SmartSteelDoc.h"
#include "SmartSteelView.h"
#include "AppDocData.h"
#include "LoginDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CSmartSteelApp

BEGIN_MESSAGE_MAP(CSmartSteelApp, OCC_3dAppEx)
	ON_COMMAND(ID_APP_ABOUT, &CSmartSteelApp::OnAppAbout)
	// Standard file based document commands
	ON_COMMAND(ID_FILE_NEW, &OCC_3dAppEx::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, &OCC_3dAppEx::OnFileOpen)
	// Standard print setup command
	ON_COMMAND(ID_FILE_PRINT_SETUP, &OCC_3dAppEx::OnFilePrintSetup)
END_MESSAGE_MAP()

#ifdef	SMART_STEEL
/******************************************************************************
    @author     humkyung
    @date       2013-08-18
    @class
    @function   SetupExceptionHandler
    @return     void
    @brief
******************************************************************************/
//static void SetupExceptionHandler()
//{
//	BT_InstallSehFilter();
//	BT_SetSupportEMail(_T("zbaekhk@gmail.com"));
//	BT_SetFlags(BTF_DETAILEDMODE | BTF_EDITMAIL);
//	BT_SetSupportServer(_T("localhost"), 9999);
//	BT_SetSupportURL(_T("http://www.solutionware.co.kr"));
//
//	// 대표적인 속성들은 다음과 같다.
//	// BTF_DETAILEDMODE : 미니덤프나 로그파일등을 압축해서 모두 전송해준다. 
//	// 정의하지 않을 경우 기본적인 정보만 전송해준다.
//     // BTF_SCREENCAPTURE : 스크린샷을 첨부한다.
//
//	// 미니덤프에 참조변수까지 남긴다.
//     BT_SetFlags( BTF_DETAILEDMODE | BTF_SCREENCAPTURE );
//
//	// Log생성
//	int g_jBT_LogSound = BT_OpenLogFile( _T("SmartSteel.log") , BTLF_TEXT );
//	BT_SetLogSizeInEntries( g_jBT_LogSound, 1024 );
//	BT_SetLogFlags( g_jBT_LogSound, BTLF_SHOWLOGLEVEL | BTLF_SHOWTIMESTAMP );
//	BT_SetLogLevel( g_jBT_LogSound, BTLL_INFO );
//	LPCTSTR pLogFileName = BT_GetLogFileName( g_jBT_LogSound );
//	BT_AddLogFile( pLogFileName );
//}
#endif

// CSmartSteelApp construction

CSmartSteelApp::CSmartSteelApp()
{

	m_bHiColorIcons = TRUE;

	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

// The one and only CSmartSteelApp object

CSmartSteelApp theApp;


/*
	@brief	return execution path
	@author	humkyung
	@date	2013.08.19
	@param
	@return	string
*/
CString GetExecPath()
{
	CString sExecPath = _T("");

	TCHAR szBuf[MAX_PATH]={'\0' ,};
	::GetModuleFileName( NULL , szBuf , MAX_PATH );
	sExecPath = szBuf;
	const int at = sExecPath.ReverseFind('\\');
	if( -1 != at )
	{
		sExecPath = sExecPath.Left(at);
	}

	return sExecPath;
}

#ifdef	SMART_STEEL
/**
	@brief  show demo splash
	@author humkyung
	@date   2014.04.25
*/
int ShowDemoSplash()
{
	CSplashScreenFx *pSplash = new CSplashScreenFx();
	if(pSplash)
	{
		pSplash->Create(GetDesktopWindow() , _T(""), 0 , CSS_FADE | CSS_CENTERSCREEN | CSS_SHADOW);
		pSplash->SetBitmap(IDB_DEMO,255,255,255);
		pSplash->Show();
		pSplash->SetActiveWindow();
	}

	return ERROR_SUCCESS;
}
#endif

// CSmartSteelApp initialization

BOOL CSmartSteelApp::InitInstance()
{
	/*_CrtSetDbgFlag( _CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF );
    _CrtSetBreakAlloc(9554);
    _CrtSetBreakAlloc(9553);
    _CrtSetBreakAlloc(9552);*/

	// InitCommonControlsEx() is required on Windows XP if an application
	// manifest specifies use of ComCtl32.dll version 6 or later to enable
	// visual styles.  Otherwise, any window creation will fail.
	INITCOMMONCONTROLSEX InitCtrls;
	InitCtrls.dwSize = sizeof(InitCtrls);
	// Set this to include all the common control classes you want to use
	// in your application.
	InitCtrls.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&InitCtrls);

	OCC_3dAppEx::InitInstance();
#if _MSC_VER < 1700
	afxAmbientActCtx = FALSE;
#endif
	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		AfxMessageBox(IDP_OLE_INIT_FAILED);
		return FALSE;
	}
	AfxEnableControlContainer();

	if ( FALSE == CSingleInstance::Create (_T("ACF4AFD4-07AA-4f8f-B5B3-27727067C5D9")))
		return FALSE ;

#ifdef	SMART_STEEL
	/// terminate if can't connect internet.
	//SetupExceptionHandler();	/// 2013.08.18 added by humkyung
#endif
	CAppDocData& docData = CAppDocData::GetInstance();

	CLoginDlg dlg;
	if(IDOK == dlg.DoModal())
	{
		docData.SetProjectName(dlg.m_rProjectNo.operator LPCTSTR());
	}else	return FALSE;
#ifdef	SMART_STEEL
	if(_T("DEMO") == docData.GetUserName()) ShowDemoSplash();
#endif
/*
	CInterpreter& python = CInterpreter::GetInstance();
	char szBuf[MAX_PATH + 1] = {'\0',};
	::GetModuleFileName( NULL , szBuf , MAX_PATH );
	char* argv[] = {szBuf};
	const char* pPath = python.startup(1 , argv);
	if(NULL != pPath)
	{
		python.AddPythonPath((GetExecPath() + "\\Python").operator LPCSTR());
		python.LoadModule("App");

		PythonCall call;
		call.call("App" , "vad2dra" , "s,s" , "a" , "b");
	}
*/

	// Standard initialization
	// If you are not using these features and wish to reduce the size
	// of your final executable, you should remove from the following
	// the specific initialization routines you do not need
	// Change the registry key under which our settings are stored
	// TODO: You should modify this string to be something appropriate
	// such as the name of your company or organization
	SetRegistryKey(_T("SmartSteel"));
	LoadStdProfileSettings(0);  // Load standard INI file options

	InitContextMenuManager();

	InitKeyboardManager();

	InitTooltipManager();
	CMFCToolTipInfo ttParams;
	ttParams.m_bVislManagerTheme = TRUE;
	theApp.GetTooltipManager()->SetTooltipParams(AFX_TOOLTIP_TYPE_ALL,
		RUNTIME_CLASS(CMFCToolTipCtrl), &ttParams);

	// Register the application's document templates.  Document templates
	//  serve as the connection between documents, frame windows and views
	CSingleDocTemplate* pDocTemplate;
	pDocTemplate = new CSingleDocTemplate(
		IDR_MAINFRAME,
		RUNTIME_CLASS(CSmartSteelDoc),
		RUNTIME_CLASS(CMainFrame),       // main SDI frame window
		RUNTIME_CLASS(CSmartSteelView));
	if (!pDocTemplate)
		return FALSE;
	AddDocTemplate(pDocTemplate);

	// Parse command line for standard shell commands, DDE, file open
	CCommandLineInfo cmdInfo;
	ParseCommandLine(cmdInfo);


	// Dispatch commands specified on the command line.  Will return FALSE if
	// app was launched with /RegServer, /Register, /Unregserver or /Unregister.
	if (!ProcessShellCommand(cmdInfo))
		return FALSE;
	
	// The one and only window has been initialized, so show and update it
	m_pMainWnd->ShowWindow(SW_SHOW);
	m_pMainWnd->UpdateWindow();
	// call DragAcceptFiles only if there's a suffix
	//  In an SDI app, this should occur after ProcessShellCommand
	return TRUE;
}



// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()

// App command to run the dialog
void CSmartSteelApp::OnAppAbout()
{
	CAboutDlg aboutDlg;
	aboutDlg.DoModal();
}

// CSmartSteelApp customization load/save methods

void CSmartSteelApp::PreLoadState()
{
	BOOL bNameValid;
	CString strName;
	bNameValid = strName.LoadString(IDS_EDIT_MENU);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_POPUP_EDIT);
	bNameValid = strName.LoadString(IDS_EXPLORER);
	ASSERT(bNameValid);
	GetContextMenuManager()->AddMenu(strName, IDR_POPUP_EXPLORER);
}

void CSmartSteelApp::LoadCustomState()
{
}

void CSmartSteelApp::SaveCustomState()
{
}

// CSmartSteelApp message handlers