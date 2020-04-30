#pragma once

#include "Resource.h"
#include <IsTools.h>

#include <gui\GridCtrl\GridCtrl.h>

// CLinearMemberShapeDlg dialog

class CLinearMemberShapeDlg : public CDialog
{
	DECLARE_DYNAMIC(CLinearMemberShapeDlg)

public:
	const int GetNewLinearMemberCount() const;
	/// add new linear member section name
	int AddNewLinearMember(const STRING_T& sSection);

	CLinearMemberShapeDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CLinearMemberShapeDlg();

// Dialog Data
	enum { IDD = IDD_LINEAR_MEMBER_SHAPE };
	CGridCtrl m_ctrlExistingMember;
	CGridCtrl m_ctrlNewMember;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
private:
	list<STRING_T> m_oNewLinearMemberList;
public:
	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedOk();
};
