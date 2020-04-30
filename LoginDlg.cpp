// LoginDlg.cpp : implementation file
//

#include "stdafx.h"
#include "LoginDlg.h"
#include "enablebuddybutton.h"
#include "AppDocData.h"
#include <gui/FolderDlg/FolderDlg.h>
#include <util/Registry.h>
#include <Tokenizer.h>

#include <util/FileTools.h>
#include <util/SplitPath.h>

#ifdef	SMART_STEEL
//#include "Socket\CommandObject.h"
//CClientSocket CLoginDlg::m_oClientSocket;
#endif

// CLoginDlg dialog

IMPLEMENT_DYNCREATE(CLoginDlg, CDialog)

CLoginDlg::CLoginDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CLoginDlg::IDD, pParent)
{
}

CLoginDlg::~CLoginDlg()
{
}

void CLoginDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);

	DDX_Control(pDX , IDC_STATIC_TITLE , m_wndStaticTitle);
	DDX_Control(pDX , IDC_STATIC_PROJECT , m_wndProjectNoStatic);
	DDX_Control(pDX , IDC_COMBO_PROJECT , m_wndProjectNoCombo);
	DDX_Text(pDX , IDC_COMBO_PROJECT , m_rProjectNo);
	/*DDX_Control(pDX , IDC_STATIC_ID , m_wndUserIDStatic);
	DDX_Control(pDX , IDC_EDIT_ID , m_wndUserID);
	*/
	DDX_Text(pDX , IDC_EDIT_ID , m_sId);
	/*
	DDX_Control(pDX , IDC_STATIC_PASSWORD , m_wndPasswordStatic);
	DDX_Control(pDX , IDC_EDIT_PASSWORD , m_wndPassword);
	DDX_Text(pDX , IDC_EDIT_FOLDER , m_rServerFolderPath);
	*/
#ifdef	SMART_STEEL
	DDX_Control(pDX , IDC_STATIC_MSG , m_wndStaticMsg);
#endif
	DDX_Control(pDX , IDOK , m_wndLoginButton);
	DDX_Control(pDX , IDCANCEL , m_wndCancelButton);
}

/**
	@brief	

	@author	humkyung
*/
BOOL CLoginDlg::OnInitDialog()
{
	CDialog::OnInitDialog();
	
	/// Set Visual Style
	CMFCVisualManagerOffice2007::SetStyle(CMFCVisualManagerOffice2007::Office2007_LunaBlue);
	CMFCVisualManager::SetDefaultManager(RUNTIME_CLASS(CMFCVisualManagerOffice2007));

	m_wndStaticTitle.SetWindowText(PRODUCT_NAME);

	auto_ptr<CFont> m_pBoldFont(new CFont);
	m_pBoldFont->CreateFont(20,0,0,0,900,0,0,0,0,0,0,ANTIALIASED_QUALITY,0, _T("Arial"));

	m_wndStaticTitle.SetColor( RGB(0, 0, 0) );
	m_wndStaticTitle.SetGradientColor( RGB(74, 84, 103) );
	m_wndStaticTitle.SetFont(m_pBoldFont.get());
	m_wndStaticTitle.SetVerticalGradient(TRUE);
	m_wndStaticTitle.SetTextAlign(DT_CENTER);
	m_wndStaticTitle.SetTextColor(RGB(255 , 255 , 255));

	m_wndProjectNoStatic.SetColor(RGB(245,245,245));
	m_wndProjectNoStatic.SetGradientColor(RGB(245,245,245));
	m_wndProjectNoStatic.SetTextColor(RGB(0, 0, 0));
#ifdef SMART_STEEL
	m_wndStaticMsg.SetGradientColor( GetSysColor(COLOR_BTNFACE) );
	m_wndStaticMsg.SetColor( GetSysColor(COLOR_BTNFACE) );
	m_wndStaticMsg.SetTextColor(RGB(0 , 0 , 0));
#endif
	CString sAppPath = CFileTools::GetCommonAppDataPath() + _T("\\") + PRODUCT_PUBLISHER + _T("\\") + PRODUCT_NAME;
	vector<CString> oFoundFolder;
	CFileTools::FindFolders(oFoundFolder , sAppPath);
	for(vector<CString>::iterator itr = oFoundFolder.begin();itr != oFoundFolder.end();++itr)
	{
		CSplitPath path(*itr);
		m_wndProjectNoCombo.AddString(path.GetFileName());
	}
	
	const CString sAppSettingFilePath = sAppPath + _T("\\") + PRODUCT_NAME + _T(".ini");
	const CString sApp(_T("General"));
	TCHAR szBuf[MAX_PATH + 1]={'\0'};
	if(GetPrivateProfileString(sApp , _T("CurPrjName") , _T("") , szBuf , MAX_PATH , sAppSettingFilePath))
	{
		m_rProjectNo = szBuf;
	}
	if(GetPrivateProfileString(sApp , _T("CurID") , _T("") , szBuf , MAX_PATH , sAppSettingFilePath))
	{
		m_sId = szBuf;
	}

	HICON hIcon = (HICON)::LoadImage( AfxGetInstanceHandle(), 
		MAKEINTRESOURCE(IDI_RESET_ACTIVATE_CODE),IMAGE_ICON,16,16, 0 );
	((CButton*)GetDlgItem(IDC_BUTTON_RESET_ACTIVATE_CODE))->SetIcon(hIcon);
#ifndef	SMART_STEEL
	GetDlgItem(IDC_BUTTON_RESET_ACTIVATE_CODE)->EnableWindow(FALSE);
#endif

	HICON hOKIcon = (HICON)::LoadImage( AfxGetInstanceHandle(), 
		MAKEINTRESOURCE(IDI_OK),
		IMAGE_ICON,
		16,16, 0 );
	m_wndLoginButton.SetIcon(hOKIcon);
	
	HICON hCancelIcon = (HICON)::LoadImage( AfxGetInstanceHandle(), 
		MAKEINTRESOURCE(IDI_CANCEL),
		IMAGE_ICON,
		16,16, 0 );
	m_wndCancelButton.SetIcon(hCancelIcon);

	UpdateData(FALSE);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

BEGIN_MESSAGE_MAP(CLoginDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CLoginDlg::OnBnClickedOk)
#ifdef	SMART_STEEL
	//ON_MESSAGE(SMARTLMS_MSG, &CLoginDlg::OnReceiveMsg)
#endif
	//ON_EN_CHANGE(IDC_EDIT_ID, &CLoginDlg::OnEnChangeEditId)
END_MESSAGE_MAP()

// CLoginDlg message handlers

/**
	@brief	
	<li> 사용자가 입력한 값들을 ini 파일로 저장한다.
	@author	BHK
	@date	?
	@param
	@return
*/
void CLoginDlg::OnBnClickedOk()
{
	/// get Project No which is Selected.
	m_wndProjectNoCombo.GetWindowText(m_rProjectNo);
	if(m_rProjectNo.IsEmpty()) return;

	CString sAppPath = CFileTools::GetCommonAppDataPath() + _T("\\") + PRODUCT_PUBLISHER + _T("\\") + PRODUCT_NAME;
	CString sPrjFolder = sAppPath + _T("\\") + m_rProjectNo;
	if(!PathFileExists(sPrjFolder))
	{
		CFileTools::CreateFolder(sPrjFolder);
		CString sAppPath;
		CFileTools::GetExecutableDirectory(sAppPath);
		::CopyFile(sAppPath + _T("Backup\\") + PRODUCT_NAME + _T(".ini") , sPrjFolder + _T("\\") + PRODUCT_NAME + _T(".ini") , TRUE);
		::CopyFile(sAppPath + _T("Backup\\") + PRODUCT_NAME + _T(".mdb") , sPrjFolder + _T("\\") + PRODUCT_NAME + _T(".mdb") , TRUE);
	}
	
	if(_T("DEMO") != m_rProjectNo)
	{
#ifdef	SMART_STEEL
		/*CString sAddress(_T("175.126.145.37"));
		DWORD dwPortNo = 2002;
		CRegistry registry;
		{
			registry.Open(HKEY_LOCAL_MACHINE , _T("Software\\TechSun\\SmartSteel\\License"));
			registry.Read(_T("IP") , sAddress);
			registry.Read(_T("Port") , dwPortNo);
			registry.Close();
		}

		if(TRUE == m_oClientSocket.Init(sAddress , dwPortNo))
		{
			TCHAR szBuf[206]={'\0',};

			m_oClientSocket.AttachWindow(this->GetSafeHwnd());
			Packet packet;
			InitializePacket(&packet);
			packet.Code = REQ_LOGIN;

			CString sString;
			GetDlgItemText(IDC_EDIT_ID , sString);
			strcpy((char*)(packet.ID) , T2CA(sString));
			GetDlgItemText(IDC_EDIT_PASSWORD , sString);
			strcpy((char*)(packet.Password) , T2CA(sString));
			strcpy((char*)(packet.AppName) , ("SmartSteel"));
			packet.Major = 1;packet.Minor=0;packet.Maintenance = 0;packet.Build = 0;
			vector<STRING_T> oResult;
			CAppDocData& docData = CAppDocData::GetInstance();
			CTokenizer<CIsFromString>::Tokenize(oResult , docData.GetFileVersion().operator LPCTSTR() , CIsFromString(_T(".")));
			if(4 == oResult.size())
			{
				packet.Major = ATOI_T(oResult[0].c_str());
				packet.Minor = ATOI_T(oResult[1].c_str());
				packet.Maintenance = ATOI_T(oResult[2].c_str());
				packet.Build = ATOI_T(oResult[3].c_str());
			}
			m_oClientSocket.Send(&packet);
		}
		else
		{
			AfxMessageBox(_T("Can't connect to server") , MB_OK);
		}*/
	CDialog::OnOK();
#else
	CDialog::OnOK();
#endif
	}
	else
	{
		CDialog::OnOK();
	}
}

/**
	@brief	ID가 수정되었을때...
	@author	HumKyung.BAEK
*/
void CLoginDlg::OnEnChangeEditId()
{
	// TODO:  If this is a RICHEDIT control, the control will not
	// send this notification unless you override the CDialog::OnInitDialog()
	// function and call CRichEditCtrl().SetEventMask()
	// with the ENM_CHANGE flag ORed into the mask.

	// TODO:  Add your control notification handler code here
	UpdateData(TRUE);
}

/**
	@brief	UI의 내용을 application setting file로 저장한다.

	@author	humkyung
*/
int CLoginDlg::SaveAppSettingFile(void)
{
	UpdateData(TRUE);

	const CString sAppSettingFilePath = CFileTools::GetCommonAppDataPath() + _T("\\") + PRODUCT_PUBLISHER + _T("\\") + PRODUCT_NAME + _T("\\") + PRODUCT_NAME + _T(".ini");
	const CString sApp(_T("General"));

	WritePrivateProfileString(sApp , _T("CurPrjName") , m_rProjectNo , sAppSettingFilePath);
	WritePrivateProfileString(sApp , _T("CurID") , m_sId , sAppSettingFilePath);

	CAppDocData& docData = CAppDocData::GetInstance();
	docData.SetUserName(m_sId);

	return ERROR_SUCCESS;
}

#ifdef	SMART_STEEL
/**
	@brief	receive message form server
	@author	humkyung
	@date	2013.06.27
*/
//LRESULT CLoginDlg::OnReceiveMsg(WPARAM wParam, LPARAM lParam)
//{
//	CCommandObject* pObj = (CCommandObject*)(wParam);
//	
//	if(GetDlgItem(IDC_BUTTON_RESET_ACTIVATE_CODE)->IsWindowEnabled())
//	{
//		if(RES_SUCCESS == pObj->m_Packet.Code)
//		{
//			//m_wndStaticMsg.SetTextColor( RGB(0,0,0) );
//			SaveAppSettingFile();
//			CDialog::OnOK();
//		}
//		else
//		{
//			m_wndStaticMsg.SetTextColor( RGB(255,0,0) );
//		}
//		m_wndStaticMsg.SetWindowText(StringHelper(pObj->m_Packet.Code));
//		m_wndLoginButton.EnableWindow(TRUE);
//	}
//	/// called from reset activate code
//	else
//	{
//		GetDlgItem(IDC_BUTTON_RESET_ACTIVATE_CODE)->EnableWindow(TRUE);
//		m_wndLoginButton.EnableWindow(TRUE);
//		m_oClientSocket.Close();
//	}
//
//	return 0L;
//}
#endif