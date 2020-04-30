#pragma once


// CSmartSteelPropertyPage dialog

namespace PropertyPage
{
class CSmartSteelPropertyPage : public CDialog
{
	DECLARE_DYNAMIC(CSmartSteelPropertyPage)

public:
	CSmartSteelPropertyPage(UINT nIDTempalte , CWnd* pParent = NULL);   // standard constructor
	virtual ~CSmartSteelPropertyPage();

// Dialog Data
	//enum { IDD = IDD_ADRAW_PROPERTYPAGE };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()

	CString m_sIniFilePath;
public:
	BOOL Create(LPCTSTR pIniFilePath , UINT nIDTemplate , CWnd* pParentWnd = 0);
	virtual int SaveData();
	virtual int LoadData();
};
};