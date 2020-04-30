// SmartSteelPropertySheet.cpp : implementation file
//

#include "stdafx.h"
#include <assert.h>
#include "SmartSteel.h"
#include "SmartSteelPropertySheet.h"


IMPLEMENT_DYNAMIC(CPropertyTreeCtrl, CTreeCtrl)

BEGIN_MESSAGE_MAP(CPropertyTreeCtrl, CTreeCtrl)
END_MESSAGE_MAP()

// CSmartSteelPropertySheet dialog

IMPLEMENT_DYNAMIC(CSmartSteelPropertySheet, CDialog)

CSmartSteelPropertySheet::CSmartSteelPropertySheet(const int& nSelectedPage , const CString& sIniFilePath , CWnd* pParent /*=NULL*/)
	: CDialog(CSmartSteelPropertySheet::IDD, pParent) , m_nSelectedPage(nSelectedPage) , m_sIniFilePath(sIniFilePath)
{
	m_pActivePage = NULL;
}

CSmartSteelPropertySheet::~CSmartSteelPropertySheet()
{
}

void CSmartSteelPropertySheet::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);

	DDX_Control(pDX , IDC_TREE_PAGE , m_wndPropertyTreeCtrl);
	DDX_Control(pDX , IDC_STATIC_PLATE_IMAGE , m_wndPlateImage);	/// 2014.09.15 added by humkyung
}

BEGIN_MESSAGE_MAP(CSmartSteelPropertySheet, CDialog)
	ON_NOTIFY(TVN_SELCHANGED, IDC_TREE_PAGE, &CSmartSteelPropertySheet::OnTvnSelchangedTreePage)
	ON_NOTIFY(NM_RCLICK, IDC_TREE_PAGE, &CSmartSteelPropertySheet::OnNMRClickTreePage)
	ON_BN_CLICKED(IDOK, &CSmartSteelPropertySheet::OnBnClickedOk)
	ON_BN_CLICKED(IDCANCEL, &CSmartSteelPropertySheet::OnBnClickedCancel)
END_MESSAGE_MAP()


// CSmartSteelPropertySheet message handlers

/**
	@brief	
	@author	humkyung
	@date	2013.08.15
*/
HTREEITEM CSmartSteelPropertySheet::AddPage(const UINT& nID , const CString& rLabel , CSmartSteelPropertyPage* pPage , const int& nImage , HTREEITEM hParent)
{
	assert(pPage && "pPage is NULL");
	HTREEITEM hItem = NULL;

	if(pPage)
	{
		if(TRUE == pPage->Create(m_sIniFilePath , nID , this))
		{
			CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
			hItem = pTreeCtrl->InsertItem(rLabel , nImage , nImage , hParent);
			m_oPageItemMap.insert(make_pair(hItem , pPage));
			pPage->ShowWindow(SW_HIDE);

			CRect rect;
			CWnd* pWnd = (CWnd*)GetDlgItem(IDC_STATIC_PAGE);
			if(pWnd)
			{
				pWnd->GetWindowRect(&rect);
				ScreenToClient(&rect);
			}
			else
			{
				rect.SetRect(0,0,100,100);
			}

			pPage->SetWindowPos(NULL , rect.left , rect.top , rect.Width() , rect.Height() , SWP_NOZORDER);
		}
	}
	else
	{
		CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
		hItem = pTreeCtrl->InsertItem(rLabel , nImage , nImage , hParent);
		m_oPageItemMap.insert(make_pair(hItem , pPage));
	}
	
	return hItem;
}

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @function   CreatePageOf
    @return     HTREEITEM
    @param      const   STRING_T&
    @param      sTitle  HTREEITEM
    @param      hParent
    @brief
******************************************************************************/
HTREEITEM CSmartSteelPropertySheet::CreatePageOf(const STRING_T& sTitle , HTREEITEM hParent)
{
	HTREEITEM hItem = NULL;

	CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
	
	if(_T("General") == sTitle)
	{
		hItem = AddPage(CPlateConfigDlg::IDD , _T("General") , &m_wndGeneralPage , 0 , hParent);
	}
	else if(_T("Steel Member Shape") == sTitle)
	{
		hItem = AddPage(CEditLinearMemberShapeDlg::IDD , _T("Steel Member Shape") , &m_wndSteelMemberShapePage , 0 , hParent);
	}
	else if(_T("Gusset Plate Shape") == sTitle)
	{
		hItem = AddPage(CEditLinearMemberShapeDlg::IDD , _T("Gusset Plate Shape") , &m_wndGussetPlateShapePage , 0 , hParent);
	}
	else if(_T("End Plate Shape") == sTitle)
	{
		hItem = AddPage(CEditLinearMemberShapeDlg::IDD , _T("End Plate Shape") , &m_wndEndPlateShapePage , 0 , hParent);
	}
	
	return hItem;
}

/**
	@brief	set active page
	@author	humkyung
	@date	2013.08.15
*/
void CSmartSteelPropertySheet::OnTvnSelchangedTreePage(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMTREEVIEW pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
	
	SetActivePage(pNMTreeView->itemNew.hItem);

	*pResult = 0;
}

/**
	@brief	
	@author	humkyung
	@date	2013.08.15
*/
void CSmartSteelPropertySheet::OnNMRClickTreePage(NMHDR *pNMHDR, LRESULT *pResult)
{
	CPoint pt , ptClient;
	::GetCursorPos(&pt);
	ptClient = pt;
	m_wndPropertyTreeCtrl.ScreenToClient(&ptClient);
	
	TVHITTESTINFO htinfo;
	htinfo.pt       = ptClient;
	htinfo.flags    = TVHT_ONITEMLABEL;
	HTREEITEM hItem = m_wndPropertyTreeCtrl.HitTest(&htinfo);
	if(NULL != hItem)
	{
	}

	*pResult = 0;
}

/**
	@brief	
	@author	humkyung
	@date	2013.08.15
*/
int CSmartSteelPropertySheet::SetActivePage(HTREEITEM hItem)
{
	assert(hItem && "hItem is NULL");

	CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
	if(hItem && pTreeCtrl)
	{
		map<HTREEITEM , CSmartSteelPropertyPage*>::iterator where = m_oPageItemMap.find(hItem);
		if((where != m_oPageItemMap.end()) && (NULL != where->second))
		{
			if(m_pActivePage != where->second)
			{
				if(NULL != m_pActivePage) m_pActivePage->ShowWindow(SW_HIDE);
				m_pActivePage = where->second;
				m_pActivePage->ShowWindow(SW_SHOWNORMAL);

				pTreeCtrl->SetItemState(hItem , TVIS_SELECTED, TVIS_SELECTED);
				pTreeCtrl->SelectItem(hItem);

				/// show plate image - 2014.09.15 added by humkyung
				CString sImageFile(_T("Empty.bmp")) , sItemText = pTreeCtrl->GetItemText(hItem);
				if(_T("Gusset Plate Shape") == sItemText)
				{
					sImageFile = _T("GussetPlate.bmp");
				}
				else if(_T("End Plate Shape") == sItemText)
				{
					sImageFile = _T("EndPlate.bmp");
				}

				m_wndPlateImage.Load(GetExecPath() + _T("\\Resource\\") + sImageFile);
				m_wndPlateImage.Draw();
				/// up to here
			}
		}
		else
		{
			HTREEITEM hChild = pTreeCtrl->GetChildItem(hItem);
			SetActivePage(hChild);
		}
	}

	return ERROR_SUCCESS;
}

BOOL CSmartSteelPropertySheet::OnInitDialog()
{
	CDialog::OnInitDialog();

	CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
	if(NULL != pTreeCtrl)
	{
		HTREEITEM hRoot = pTreeCtrl->InsertItem(_T("Options"));
		if(NULL != hRoot)
		{
			vector<HTREEITEM> oTreeItemList;
			oTreeItemList.push_back(CreatePageOf(_T("General") , hRoot));
			oTreeItemList.push_back(CreatePageOf(_T("Steel Member Shape") , hRoot));
			oTreeItemList.push_back(CreatePageOf(_T("Gusset Plate Shape") , hRoot));
			oTreeItemList.push_back(CreatePageOf(_T("End Plate Shape") , hRoot));
			pTreeCtrl->Expand(hRoot , TVE_EXPAND);

			pTreeCtrl->SetFocus();
			SetActivePage(oTreeItemList[m_nSelectedPage]);
		}
	}

	return FALSE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

/**
	@brief	
	@author	humkyung
	@date	2013.08.15
*/
void CSmartSteelPropertySheet::OnBnClickedOk()
{
	CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
	for(map<HTREEITEM , CSmartSteelPropertyPage*>::iterator itr = m_oPageItemMap.begin();itr != m_oPageItemMap.end();++itr)
	{
		const CString sItemText = pTreeCtrl->GetItemText(itr->first);

		if(NULL == itr->second) continue;
		if(itr->second->IsKindOf(RUNTIME_CLASS(CSmartSteelPropertyPage)))
		{
			CSmartSteelPropertyPage* pPage = (CSmartSteelPropertyPage*)(itr->second);
			if(pPage) pPage->SaveData();
		}
	}

	OnOK();
}

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @function   OnBnClickedCancel
    @return     void
    @brief
******************************************************************************/
void CSmartSteelPropertySheet::OnBnClickedCancel()
{
	CTreeCtrl* pTreeCtrl = (CTreeCtrl*)GetDlgItem(IDC_TREE_PAGE);
	for(map<HTREEITEM , CSmartSteelPropertyPage*>::iterator itr = m_oPageItemMap.begin();itr != m_oPageItemMap.end();++itr)
	{
		const CString sItemText = pTreeCtrl->GetItemText(itr->first);

		if(NULL == itr->second) continue;
		if(itr->second->IsKindOf(RUNTIME_CLASS(CSmartSteelPropertyPage)))
		{
		}
	}

	OnCancel();
}

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @function   PreTranslateMessage
    @return     BOOL
    @param      MSG*    pMsg
    @brief
******************************************************************************/
BOOL CSmartSteelPropertySheet::PreTranslateMessage(MSG* pMsg)
{
	if(pMsg->message == WM_KEYDOWN)
	{
		if(pMsg->wParam == VK_ESCAPE) return TRUE;
		if(pMsg->wParam == VK_RETURN) return TRUE;
	}

	return CDialog::PreTranslateMessage(pMsg);
}