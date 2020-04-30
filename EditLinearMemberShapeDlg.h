#pragma once
#include <gui\GridCtrl\GridCtrl.h>

#include "SmartSteelPropertyPage.h"

// CEditLinearMemberShapeDlg dialog

namespace PropertyPage
{
class CEditLinearMemberShapeDlg : public CSmartSteelPropertyPage
{
	DECLARE_DYNAMIC(CEditLinearMemberShapeDlg)

	typedef enum
	{
		NONE_ITEM		= 0x00,
		NEW_ITEM		= 0x01,
		DELETED_ITEM	= 0x02,
		MODIFIED_ITEM	= 0x04,
	}ItemStatus;
public:
	CEditLinearMemberShapeDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CEditLinearMemberShapeDlg();

// Dialog Data
	enum { IDD = IDD_EDIT_LINEAR_MEMBER_SHAPE };
	CGridCtrl m_ctrlLinearMember;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
public:
	int SaveData();

	virtual BOOL OnInitDialog();
	afx_msg void OnBnClickedNewLinearMember();
	afx_msg void OnBnClickedDeleteLinearMember();
protected:
	virtual BOOL OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult);
};
};