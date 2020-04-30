#pragma once
/////////////////////////////////////////////////////////////////////////////
// CAboutDlgStd dialog used for App About

class CAboutDlgStd : public CDialog
{
public:
	CAboutDlgStd();
	BOOL OnInitDialog();

	// Dialog Data
	//{{AFX_DATA(CAboutDlgStd)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlgStd)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL


	// Implementation
protected:
	//{{AFX_MSG(CAboutDlgStd)
	// No message handlers
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};