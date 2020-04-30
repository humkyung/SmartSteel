#pragma once
#include <gui\GridCtrl\GridCtrl.h>

#include "SmartSteelPropertyPage.h"

// CEndPlateShapeDlg dialog

namespace PropertyPage
{
class CEndPlateShapeDlg : public CSmartSteelPropertyPage
{
	DECLARE_DYNAMIC(CEndPlateShapeDlg)

	typedef enum
	{
		NONE_ITEM		= 0x00,
		NEW_ITEM		= 0x01,
		DELETED_ITEM	= 0x02,
		MODIFIED_ITEM	= 0x04,
	}ItemStatus;
public:
	CEndPlateShapeDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CEndPlateShapeDlg();

// Dialog Data
	enum { IDD = IDD_END_PLATE_SHAPE };
	CGridCtrl m_ctrlSteelPlate;
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