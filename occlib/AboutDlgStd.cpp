#include "StdAfx.h"
#include "AboutDlgStd.h"

CAboutDlgStd::CAboutDlgStd() : CDialog(CAboutDlgStd::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlgStd)
	//}}AFX_DATA_INIT
}

void CAboutDlgStd::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlgStd)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlgStd, CDialog)
	//{{AFX_MSG_MAP(CAboutDlgStd)
	// No message handlers
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

BOOL CAboutDlgStd::OnInitDialog()
{
	/*CWnd* Title = GetDlgItem(IDC_ABOUTBOX_TITLE);

	CString About = "About ";
	CString Sample = "Sample ";
	CString SampleName = ((OCC_BaseApp*)AfxGetApp())->GetSampleName();
	CString Cascade = ", Open CASCADE Technology ";
	CString Version = OCC_VERSION_STRING;

	CString strTitle = Sample + SampleName + Cascade + Version;
	CString dlgTitle = About + SampleName;

	Title->SetWindowText(strTitle);
	SetWindowText(dlgTitle);

	CenterWindow();*/

	return TRUE;
}