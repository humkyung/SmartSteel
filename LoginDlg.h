#pragma once
//#include <gui/FilterEdit/BoostBaseFilterEdit.h>
#include <gui/GradientStatic.h>
#include "Resource.h"

// CLoginDlg dialog

class CLoginDlg : public CDialog
{
	DECLARE_DYNCREATE(CLoginDlg)

public:
	CLoginDlg(CWnd* pParent = NULL);   // standard constructor
	virtual ~CLoginDlg();
// Overrides

// Dialog Data
	enum { IDD = IDD_LOGIN };

	CGradientStatic			m_wndStaticTitle;
	CGradientStatic			m_wndProjectNoStatic;
	CComboBox				m_wndProjectNoCombo;
	CString					m_rProjectNo;
	CString					m_sId;
	/*CGradientStatic			m_wndUserIDStatic;
	CBoostBaseFilterEdit    m_wndUserID;
	CGradientStatic			m_wndPasswordStatic;
	CBoostBaseFilterEdit	m_wndPassword;
	*/
#ifdef	SMART_STEEL
	CGradientStatic			m_wndStaticMsg;
#endif

	CMFCButton		m_wndLoginButton;
	CMFCButton		m_wndCancelButton;
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	virtual BOOL OnInitDialog();
	afx_msg LRESULT OnReceiveMsg(WPARAM wParam, LPARAM lParam);  /// Ãß°¡
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
public:
	afx_msg void OnEnChangeEditId();
private:
	int SaveAppSettingFile(void);
	int DisplayMessage(const string& rMsg , COLORREF bgColor , COLORREF fgColor);
private:
#ifdef	SMART_STEEL
	//static CClientSocket m_oClientSocket;
#endif
};
