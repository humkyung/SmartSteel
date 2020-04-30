// OCC_BaseAppEx.cpp: implementation of the OCC_BaseApp class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Resource.h"
#include "AboutDlgStd.h"

#include "OCC_BaseAppEx.h"
#include <Standard_Version.hxx>

/////////////////////////////////////////////////////////////////////////////
// OCC_BaseApp

BEGIN_MESSAGE_MAP(OCC_BaseAppEx, CWinAppEx)
	//{{AFX_MSG_MAP(OCC_BaseApp)
	//ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
	// NOTE - the ClassWizard will add and remove mapping macros here.
	//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
	// Standard file based document commands
	ON_COMMAND(ID_FILE_NEW, CWinAppEx::OnFileNew)
	ON_COMMAND(ID_FILE_OPEN, CWinAppEx::OnFileOpen)
	ON_COMMAND(ID_EDIT_REFRESH, &OCC_BaseAppEx::OnEditRefresh)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// OCC_BaseApp construction

OCC_BaseAppEx::OCC_BaseAppEx()
{
	SampleName = _T("");
	SetSamplePath(NULL);
}


void OCC_BaseAppEx::SetSamplePath(LPCTSTR aPath)
{
	TCHAR AbsoluteExecutableFileName[MAX_PATH+1];
	HMODULE hModule = GetModuleHandle(NULL);
	GetModuleFileName(hModule, AbsoluteExecutableFileName, MAX_PATH);

	SamplePath = CString(AbsoluteExecutableFileName);
	int index = SamplePath.ReverseFind('\\');
	SamplePath.Delete(index+1, SamplePath.GetLength() - index - 1);
	if (aPath == NULL)
		SamplePath += _T("..");
	else{
		CString aCInitialDir(aPath);
		SamplePath += _T("..\\") + aCInitialDir;
	}
}

// App command to run the dialog
//void OCC_BaseAppEx::OnAppAbout()
//{
//	CAboutDlgStd aboutDlg;
//	aboutDlg.DoModal();
//}

LPCTSTR OCC_BaseAppEx::GetSampleName()
{
	return SampleName;
}

LPCTSTR OCC_BaseAppEx::GetInitDataDir()
{
	return (LPCTSTR) SamplePath;
}

void OCC_BaseAppEx::SetSampleName(LPCTSTR Name)
{
	SampleName = Name;
}

void OCC_BaseAppEx::OnEditRefresh()
{
	// TODO: Add your command handler code here
}


/////////////////////////////////////////////////////////////////////////////
// CAboutDlgStd dialog used for App About
