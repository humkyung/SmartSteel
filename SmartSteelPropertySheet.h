#pragma once

#include <PictureWndEx.h>
#include "SmartSteelPropertyPage.h"
#include "PlateConfigDlg.h"

#include "EditLinearMemberShapeDlg.h"
#include "GussetPlateShapeDlg.h"
#include "EndPlateShapeDlg.h"

#include <map>
using namespace std;
using namespace PropertyPage;

class CPropertyTreeCtrl : public CTreeCtrl
{
	DECLARE_DYNAMIC(CPropertyTreeCtrl)
public:
	CPropertyTreeCtrl(){}
	~CPropertyTreeCtrl(){}
protected:
	DECLARE_MESSAGE_MAP()
};

// CSmartSteelPropertySheet dialog

class CSmartSteelPropertySheet : public CDialog
{
	DECLARE_DYNAMIC(CSmartSteelPropertySheet)

public:
	CSmartSteelPropertySheet(const int& nSelectedPage , const CString& , CWnd* pParent = NULL);   // standard constructor
	virtual ~CSmartSteelPropertySheet();

	// Dialog Data
	enum { IDD = IDD_PROPERTY_SHEET };
	
	CPropertyTreeCtrl m_wndPropertyTreeCtrl;
	CPictureWndEx  m_wndPlateImage;	/// 2014.09.15 added by humkyung

	CPlateConfigDlg m_wndGeneralPage;
	CEditLinearMemberShapeDlg m_wndSteelMemberShapePage;
	CGussetPlateShapeDlg m_wndGussetPlateShapePage;
	CEndPlateShapeDlg m_wndEndPlateShapePage;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnTvnSelchangedTreePage(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnNMRClickTreePage(NMHDR *pNMHDR, LRESULT *pResult);
	virtual BOOL PreTranslateMessage(MSG* pMsg);
private:
	int SetActivePage(HTREEITEM hItem);
	HTREEITEM AddPage(const UINT& nID , const CString& rLabel , CSmartSteelPropertyPage* pPage , const int& nImage , HTREEITEM hParent);
	HTREEITEM CreatePageOf(const STRING_T& sTitle , HTREEITEM hParent);
private:
	CString m_sIniFilePath;
	int m_nSelectedPage;
	CSmartSteelPropertyPage* m_pActivePage;
	map<HTREEITEM , CSmartSteelPropertyPage*> m_oPageItemMap;
public:
	virtual BOOL OnInitDialog();
};
