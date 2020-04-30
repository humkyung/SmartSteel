// EndPlateShapeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "SmartSteel.h"
#include "EndPlateShapeDlg.h"
#include <ado/ADODB.h>
#include <FileTools.h>

#include <GridCtrl/GridCellCombo.h>

#include "AppDocData.h"
// CEndPlateShapeDlg dialog

using namespace PropertyPage;

IMPLEMENT_DYNAMIC(CEndPlateShapeDlg, CSmartSteelPropertyPage)

CEndPlateShapeDlg::CEndPlateShapeDlg(CWnd* pParent /*=NULL*/)
	: CSmartSteelPropertyPage(CEndPlateShapeDlg::IDD, pParent)
{

}

CEndPlateShapeDlg::~CEndPlateShapeDlg()
{
}

void CEndPlateShapeDlg::DoDataExchange(CDataExchange* pDX)
{
	CSmartSteelPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CEndPlateShapeDlg, CSmartSteelPropertyPage)
	ON_BN_CLICKED(IDC_NEW_LINEAR_MEMBER, &CEndPlateShapeDlg::OnBnClickedNewLinearMember)
	ON_BN_CLICKED(IDC_DELETE_LINEAR_MEMBER, &CEndPlateShapeDlg::OnBnClickedDeleteLinearMember)
END_MESSAGE_MAP()


// CEndPlateShapeDlg message handlers

BOOL CEndPlateShapeDlg::OnInitDialog()
{
	CSmartSteelPropertyPage::OnInitDialog();

	CRect rect , tmpRect;

	GetDlgItem(IDC_STATIC_LINEAR_MEMBER_SHAPE)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	GetDlgItem(IDC_DELETE_LINEAR_MEMBER)->GetWindowRect(&tmpRect);
	ScreenToClient(&tmpRect);
	rect.top = tmpRect.bottom;
	rect.DeflateRect(10 , 10);
	if(TRUE == m_ctrlSteelPlate.Create(rect ,this , 0x101))
	{
		m_ctrlSteelPlate.SetFixedColumnCount(1);
		m_ctrlSteelPlate.SetFixedRowCount(1);
		m_ctrlSteelPlate.SetRowCount(CSmartSteelDoc::m_oEndPlateParamMap.size() + 1);
		m_ctrlSteelPlate.SetColumnCount(4);

		m_ctrlSteelPlate.SetItemText(0 , 0 , _T("Section Name"));
		m_ctrlSteelPlate.SetColumnWidth(0 , 120);
		m_ctrlSteelPlate.SetItemText(0 , 1 , _T("K"));
		m_ctrlSteelPlate.SetColumnWidth(1 , 80);
		m_ctrlSteelPlate.SetItemText(0 , 2 , _T("M"));
		m_ctrlSteelPlate.SetColumnWidth(2 , 70);
		m_ctrlSteelPlate.SetItemText(0 , 3 , _T("T"));
		m_ctrlSteelPlate.SetColumnWidth(3 , 70);

		int i = 1;
		OSTRINGSTREAM_T oss;
		oss.precision( 3 );
		for(map<STRING_T , CEndPlate::Param* >::iterator itr = CSmartSteelDoc::m_oEndPlateParamMap.begin();itr != CSmartSteelDoc::m_oEndPlateParamMap.end();++itr)
		{
			m_ctrlSteelPlate.SetItemText(i , 0 , itr->first.c_str());
			m_ctrlSteelPlate.SetItemData(i , 0 , CEndPlateShapeDlg::NONE_ITEM);
			oss.str(_T(""));
			oss << itr->second->K;
			m_ctrlSteelPlate.SetItemText(i , 1 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->M;
			m_ctrlSteelPlate.SetItemText(i , 2 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->T;
			m_ctrlSteelPlate.SetItemText(i , 3 , oss.str().c_str());
			++i;
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

/**
	@brief	add a new linear member shape parameter
	@author	humkyung
	@date	2013.08.13
*/
void CEndPlateShapeDlg::OnBnClickedNewLinearMember()
{
	CString sSectionName;
	GetDlgItemText(IDC_EDIT_SECTION_NAME,sSectionName);
	if(sSectionName.IsEmpty())
	{
		AfxMessageBox(_T("Section Name is empty!!!") , MB_OK | MB_ICONWARNING);
		return;
	}
	if(CSmartSteelDoc::m_oEndPlateParamMap.end() != CSmartSteelDoc::m_oEndPlateParamMap.find(sSectionName.operator LPCTSTR()))
	{
		AfxMessageBox(_T("Section Name is already existing\nPlease, input another section name!!!") , MB_OK | MB_ICONWARNING);
		return;
	}

	if(TRUE == m_ctrlSteelPlate.SetRowCount( m_ctrlSteelPlate.GetRowCount() + 1 ))
	{
		m_ctrlSteelPlate.SetItemText(m_ctrlSteelPlate.GetRowCount() - 1, 0 , sSectionName);
		m_ctrlSteelPlate.SetItemData(m_ctrlSteelPlate.GetRowCount() - 1 , 0 , CEndPlateShapeDlg::NEW_ITEM);
	}
}

/**
	@brief	delete selected linear member shape
	@author	humkyung
	@date	2013.08.12
*/
void CEndPlateShapeDlg::OnBnClickedDeleteLinearMember()
{
	CAppDocData& docData = CAppDocData::GetInstance();
	
	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		CCellRange range = m_ctrlSteelPlate.GetSelectedCellRange();
		for(int i = range.GetMaxRow();i >= range.GetMinRow();--i)
		{
			const CString sSectionName = m_ctrlSteelPlate.GetItemText(i , 0);
			if(CEndPlateShapeDlg::NEW_ITEM != m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// delete item from database
				CString sSql;
				sSql.Format(_T("DELETE FROM END_PLATE WHERE SectionName='%s'") , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
			}
			map<STRING_T , CEndPlate::Param* >::iterator where = CSmartSteelDoc::m_oEndPlateParamMap.find(sSectionName.operator LPCTSTR());
			if(where != CSmartSteelDoc::m_oEndPlateParamMap.end())
			{
				SAFE_DELETE(where->second);
				CSmartSteelDoc::m_oEndPlateParamMap.erase(where);
			}

			m_ctrlSteelPlate.DeleteRow(i);
		}
		m_ctrlSteelPlate.RedrawWindow();

		adodb.DBDisConnect();
	}
}

/**
	@brief	save new or modified stee shape parameter
	@author	humkyung
	@date	2013.08.13
*/
int CEndPlateShapeDlg::SaveData()
{
	CAppDocData& docData = CAppDocData::GetInstance();
	
	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		CCellRange range = m_ctrlSteelPlate.GetSelectedCellRange();
		for(int i = 1;i < m_ctrlSteelPlate.GetRowCount();++i)
		{
			const CString sSectionName = m_ctrlSteelPlate.GetItemText(i , 0);
			const CString sK = m_ctrlSteelPlate.GetItemText(i , 1);
			const CString sM = m_ctrlSteelPlate.GetItemText(i , 2);
			const CString sT = m_ctrlSteelPlate.GetItemText(i , 3);
			if(sSectionName.IsEmpty() || sK.IsEmpty() || sM.IsEmpty() || sT.IsEmpty()) continue;

			CString sSql;
			if(CEndPlateShapeDlg::NEW_ITEM == m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// insert item to database
				sSql.Format(_T("INSERT INTO END_PLATE(SectionName,K,M,T) VALUES('%s',%s,%s,%s)") , sSectionName , sK , sM , sT);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				
				CEndPlate::Param* pParam = new CEndPlate::Param;
				if(NULL != pParam)
				{
					pParam->K = ATOF_T(sK.operator LPCTSTR());
					pParam->M = ATOF_T(sM.operator LPCTSTR());
					pParam->T = ATOF_T(sT.operator LPCTSTR());
					CSmartSteelDoc::m_oEndPlateParamMap.insert(make_pair(sSectionName.operator LPCTSTR() , pParam));
				}
			}
			else if(CEndPlateShapeDlg::MODIFIED_ITEM == m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// update item
				sSql.Format(_T("UPDATE END_PLATE SET K=%s,M=%s,T=%s WHERE SectionName='%s'") , sK , sM , sT , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				map<STRING_T , CEndPlate::Param* >::iterator where = CSmartSteelDoc::m_oEndPlateParamMap.find(sSectionName.operator LPCTSTR());
				if(where != CSmartSteelDoc::m_oEndPlateParamMap.end())
				{
					where->second->K = ATOF_T(sK.operator LPCTSTR());
					where->second->M = ATOF_T(sM.operator LPCTSTR());
					where->second->T = ATOF_T(sT.operator LPCTSTR());
				}
			}
		}
		
		adodb.DBDisConnect();
	}

	return ERROR_SUCCESS;
}

/**
	@brief	notify from child
	@author	humkyung
	@date	2013.08.13
*/
BOOL CEndPlateShapeDlg::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	if(wParam == m_ctrlSteelPlate.GetDlgCtrlID())
	{
		NM_GRIDVIEW* pGridView = (NM_GRIDVIEW*)(lParam);

		if(GVN_ENDLABELEDIT == pGridView->hdr.code)
		{
			CGridCellBase* pCell = m_ctrlSteelPlate.GetCell(pGridView->iRow, pGridView->iColumn);
			if(pCell)
			{
				pCell->SetTextClr(RGB(255,0,0));
				pCell->EndEdit();
			}
			if(CEndPlateShapeDlg::NEW_ITEM != m_ctrlSteelPlate.GetItemData(pGridView->iRow , pGridView->iColumn))
			{
				m_ctrlSteelPlate.SetItemData(pGridView->iRow , pGridView->iColumn , CEndPlateShapeDlg::MODIFIED_ITEM);
			}
			return TRUE;
		}
	}

	return CSmartSteelPropertyPage::OnNotify(wParam, lParam, pResult);
}
