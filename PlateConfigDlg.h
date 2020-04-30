#pragma once

#include "StdAfx.h"
#include "SmartSteelPropertyPage.h"

namespace PropertyPage
{
// CPlateConfigDlg dialog

class CPlateConfigDlg : public CSmartSteelPropertyPage
{
	DECLARE_DYNAMIC(CPlateConfigDlg)

public:
	CPlateConfigDlg(const CString sIniFilePath = _T("") , CWnd* pParent = NULL);   // standard constructor
	virtual ~CPlateConfigDlg();

// Dialog Data
	enum { IDD = IDD_PLATE_CONFIG };

	PlateCfg m_oPlateCfg;
	BOOL m_bGenerateForWebTypeBrace;
	BOOL m_bGenerateEndPlateDependOnBeamLength;
	CString m_sClass , m_sGrade , m_sGussetPlateDisplayColor , m_sEndPlateDisplayColor;
	CString m_sUnit;	/// 2014.02.08 added by humkyung
	CString m_sMaxEdgeLengthToMerge;	/// 2014.02.14 added by humkyung
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	int SaveData();

	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
	afx_msg void OnBnClickedEditLinearMemberShape();
};
};