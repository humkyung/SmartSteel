// GussetPlateShapeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "SmartSteel.h"
#include "GussetPlateShapeDlg.h"
#include <ado/ADODB.h>
#include <FileTools.h>

#include <GridCtrl/GridCellCombo.h>

#include "AppDocData.h"
// CGussetPlateShapeDlg dialog

using namespace PropertyPage;

IMPLEMENT_DYNAMIC(CGussetPlateShapeDlg, CSmartSteelPropertyPage)

CGussetPlateShapeDlg::CGussetPlateShapeDlg(CWnd* pParent /*=NULL*/)
	: CSmartSteelPropertyPage(CGussetPlateShapeDlg::IDD, pParent)
{

}

CGussetPlateShapeDlg::~CGussetPlateShapeDlg()
{
}

void CGussetPlateShapeDlg::DoDataExchange(CDataExchange* pDX)
{
	CSmartSteelPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CGussetPlateShapeDlg, CSmartSteelPropertyPage)
	ON_BN_CLICKED(IDC_NEW_LINEAR_MEMBER, &CGussetPlateShapeDlg::OnBnClickedNewLinearMember)
	ON_BN_CLICKED(IDC_DELETE_LINEAR_MEMBER, &CGussetPlateShapeDlg::OnBnClickedDeleteLinearMember)
END_MESSAGE_MAP()


// CGussetPlateShapeDlg message handlers

BOOL CGussetPlateShapeDlg::OnInitDialog()
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
		m_ctrlSteelPlate.SetRowCount(CSmartSteelDoc::m_oGussetPlateParamMap.size() + 1);
		m_ctrlSteelPlate.SetColumnCount(7);

		m_ctrlSteelPlate.SetItemText(0 , 0 , _T("Section Name"));
		m_ctrlSteelPlate.SetColumnWidth(0 , 120);
		m_ctrlSteelPlate.SetItemText(0 , 1 , _T("A"));
		m_ctrlSteelPlate.SetColumnWidth(1 , 80);
		m_ctrlSteelPlate.SetItemText(0 , 2 , _T("B"));
		m_ctrlSteelPlate.SetColumnWidth(2 , 60);
		m_ctrlSteelPlate.SetItemText(0 , 3 , _T("E"));
		m_ctrlSteelPlate.SetColumnWidth(3 , 60);
		m_ctrlSteelPlate.SetItemText(0 , 4 , _T("N"));
		m_ctrlSteelPlate.SetColumnWidth(4 , 60);
		m_ctrlSteelPlate.SetItemText(0 , 5 , _T("P"));
		m_ctrlSteelPlate.SetColumnWidth(5 , 60);
		m_ctrlSteelPlate.SetItemText(0 , 6 , _T("T"));
		m_ctrlSteelPlate.SetColumnWidth(6 , 60);

		int i = 1;
		OSTRINGSTREAM_T oss;
		oss.precision( 3 );
		for(map<STRING_T , CGussetPlate::Param* >::iterator itr = CSmartSteelDoc::m_oGussetPlateParamMap.begin();itr != CSmartSteelDoc::m_oGussetPlateParamMap.end();++itr)
		{
			m_ctrlSteelPlate.SetItemText(i , 0 , itr->first.c_str());
			m_ctrlSteelPlate.SetItemData(i , 0 , CGussetPlateShapeDlg::NONE_ITEM);

			oss.str(_T(""));
			oss << itr->second->A;
			m_ctrlSteelPlate.SetItemText(i , 1 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->B;
			m_ctrlSteelPlate.SetItemText(i , 2 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->E;
			m_ctrlSteelPlate.SetItemText(i , 3 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->N;
			m_ctrlSteelPlate.SetItemText(i , 4 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->P;
			m_ctrlSteelPlate.SetItemText(i , 5 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->T;
			m_ctrlSteelPlate.SetItemText(i , 6 , oss.str().c_str());
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
void CGussetPlateShapeDlg::OnBnClickedNewLinearMember()
{
	CString sSectionName;
	GetDlgItemText(IDC_EDIT_SECTION_NAME,sSectionName);
	if(sSectionName.IsEmpty())
	{
		AfxMessageBox(_T("Section Name is empty!!!") , MB_OK | MB_ICONWARNING);
		return;
	}
	if(CSmartSteelDoc::m_oGussetPlateParamMap.end() != CSmartSteelDoc::m_oGussetPlateParamMap.find(sSectionName.operator LPCTSTR()))
	{
		AfxMessageBox(_T("Section Name is already existing\nPlease, input another section name!!!") , MB_OK | MB_ICONWARNING);
		return;
	}

	if(TRUE == m_ctrlSteelPlate.SetRowCount( m_ctrlSteelPlate.GetRowCount() + 1 ))
	{
		m_ctrlSteelPlate.SetItemText(m_ctrlSteelPlate.GetRowCount() - 1 , 0 , sSectionName);
		m_ctrlSteelPlate.SetItemData(m_ctrlSteelPlate.GetRowCount() - 1 , 0 , CGussetPlateShapeDlg::NEW_ITEM);
	}
}

/**
	@brief	delete selected linear member shape
	@author	humkyung
	@date	2013.08.12
*/
void CGussetPlateShapeDlg::OnBnClickedDeleteLinearMember()
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
			if(CGussetPlateShapeDlg::NEW_ITEM != m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// delete item from database
				CString sSql;
				sSql.Format(_T("DELETE FROM GUSSET_PLATE WHERE SectionName='%s'") , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
			}
			map<STRING_T , CGussetPlate::Param* >::iterator where = CSmartSteelDoc::m_oGussetPlateParamMap.find(sSectionName.operator LPCTSTR());
			if(where != CSmartSteelDoc::m_oGussetPlateParamMap.end())
			{
				SAFE_DELETE(where->second);
				CSmartSteelDoc::m_oGussetPlateParamMap.erase(where);
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
int CGussetPlateShapeDlg::SaveData()
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
			const CString sA = m_ctrlSteelPlate.GetItemText(i , 1);
			const CString sB = m_ctrlSteelPlate.GetItemText(i , 2);
			const CString sE  = m_ctrlSteelPlate.GetItemText(i , 3);
			const CString sN = m_ctrlSteelPlate.GetItemText(i , 4);
			const CString sP = m_ctrlSteelPlate.GetItemText(i , 5);
			const CString sT = m_ctrlSteelPlate.GetItemText(i , 6);
			if(sSectionName.IsEmpty() || sA.IsEmpty() || sB.IsEmpty() || sE.IsEmpty() || sN.IsEmpty() || sP.IsEmpty() || sT.IsEmpty()) continue;

			CString sSql;
			if(CGussetPlateShapeDlg::NEW_ITEM == m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// insert item to database
				sSql.Format(_T("INSERT INTO GUSSET_PLATE(SectionName,A,B,E,N,P,T) VALUES('%s',%s,%s,%s,%s,%s,%s)") , sSectionName , sA , sB , sE , sN , sP , sT);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				
				CGussetPlate::Param* pParam = new CGussetPlate::Param;
				if(NULL != pParam)
				{
					pParam->A = ATOF_T(sA.operator LPCTSTR());
					pParam->B = ATOF_T(sB.operator LPCTSTR());
					pParam->E = ATOF_T(sE.operator LPCTSTR());
					pParam->N = ATOF_T(sN.operator LPCTSTR());
					pParam->P = ATOF_T(sP.operator LPCTSTR());
					pParam->T = ATOF_T(sT.operator LPCTSTR());
					CSmartSteelDoc::m_oGussetPlateParamMap.insert(make_pair(sSectionName.operator LPCTSTR() , pParam));
				}
			}
			else if(CGussetPlateShapeDlg::MODIFIED_ITEM == m_ctrlSteelPlate.GetItemData(i , 0))
			{
				/// update item
				sSql.Format(_T("UPDATE GUSSET_PLATE SET A=%s,B=%s,E=%s,N=%s,P=%s,T=%s WHERE SectionName='%s'") , sA , sB , sE , sN , sP , sT , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				map<STRING_T , CGussetPlate::Param* >::iterator where = CSmartSteelDoc::m_oGussetPlateParamMap.find(sSectionName.operator LPCTSTR());
				if(where != CSmartSteelDoc::m_oGussetPlateParamMap.end())
				{
					where->second->A = ATOF_T(sA.operator LPCTSTR());
					where->second->B = ATOF_T(sB.operator LPCTSTR());
					where->second->E = ATOF_T(sE.operator LPCTSTR());
					where->second->N = ATOF_T(sB.operator LPCTSTR());
					where->second->P = ATOF_T(sP.operator LPCTSTR());
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
BOOL CGussetPlateShapeDlg::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
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
			if(CGussetPlateShapeDlg::NEW_ITEM != m_ctrlSteelPlate.GetItemData(pGridView->iRow , pGridView->iColumn))
			{
				m_ctrlSteelPlate.SetItemData(pGridView->iRow , pGridView->iColumn , CGussetPlateShapeDlg::MODIFIED_ITEM);
			}
			return TRUE;
		}
	}

	return CSmartSteelPropertyPage::OnNotify(wParam, lParam, pResult);
}
