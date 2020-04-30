// EditLinearMemberShapeDlg.cpp : implementation file
//

#include "stdafx.h"
#include "SmartSteel.h"
#include "EditLinearMemberShapeDlg.h"
#include <ado/ADODB.h>
#include <FileTools.h>
#include "AppDocData.h"

#include <GridCtrl/GridCellCombo.h>
// CEditLinearMemberShapeDlg dialog

using namespace PropertyPage;

IMPLEMENT_DYNAMIC(CEditLinearMemberShapeDlg, CSmartSteelPropertyPage)

CEditLinearMemberShapeDlg::CEditLinearMemberShapeDlg(CWnd* pParent /*=NULL*/)
	: CSmartSteelPropertyPage(CEditLinearMemberShapeDlg::IDD, pParent)
{

}

CEditLinearMemberShapeDlg::~CEditLinearMemberShapeDlg()
{
}

void CEditLinearMemberShapeDlg::DoDataExchange(CDataExchange* pDX)
{
	CSmartSteelPropertyPage::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CEditLinearMemberShapeDlg, CSmartSteelPropertyPage)
	ON_BN_CLICKED(IDC_NEW_LINEAR_MEMBER, &CEditLinearMemberShapeDlg::OnBnClickedNewLinearMember)
	ON_BN_CLICKED(IDC_DELETE_LINEAR_MEMBER, &CEditLinearMemberShapeDlg::OnBnClickedDeleteLinearMember)
END_MESSAGE_MAP()


// CEditLinearMemberShapeDlg message handlers

BOOL CEditLinearMemberShapeDlg::OnInitDialog()
{
	CSmartSteelPropertyPage::OnInitDialog();

	CAppDocData& docData = CAppDocData::GetInstance();
	const double dMultiplier = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1000.0) : 1.;

	CRect rect , tmpRect;

	GetDlgItem(IDC_STATIC_LINEAR_MEMBER_SHAPE)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	GetDlgItem(IDC_DELETE_LINEAR_MEMBER)->GetWindowRect(&tmpRect);
	ScreenToClient(&tmpRect);
	rect.top = tmpRect.bottom;
	rect.DeflateRect(10 , 10);
	if(TRUE == m_ctrlLinearMember.Create(rect ,this , 0x101))
	{
		m_ctrlLinearMember.SetFixedColumnCount(1);
		m_ctrlLinearMember.SetFixedRowCount(1);
		m_ctrlLinearMember.SetRowCount(CSmartSteelDoc::m_oShapeParamMap.size() + 1);
		m_ctrlLinearMember.SetColumnCount(6);

		m_ctrlLinearMember.SetItemText(0 , 0 , _T("Section Name"));
		m_ctrlLinearMember.SetColumnWidth(0 , 120);
		m_ctrlLinearMember.SetItemText(0 , 1 , _T("Shape"));
		m_ctrlLinearMember.SetColumnWidth(1 , 80);
		m_ctrlLinearMember.SetItemText(0 , 2 , _T("H"));
		m_ctrlLinearMember.SetColumnWidth(2 , 70);
		m_ctrlLinearMember.SetItemText(0 , 3 , _T("B"));
		m_ctrlLinearMember.SetColumnWidth(3 , 70);
		m_ctrlLinearMember.SetItemText(0 , 4 , _T("T1"));
		m_ctrlLinearMember.SetColumnWidth(4 , 70);
		m_ctrlLinearMember.SetItemText(0 , 5 , _T("T2"));
		m_ctrlLinearMember.SetColumnWidth(5 , 70);

		int i = 1;
		OSTRINGSTREAM_T oss;
		oss.precision( 3 );
		for(map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator itr = CSmartSteelDoc::m_oShapeParamMap.begin();itr != CSmartSteelDoc::m_oShapeParamMap.end();++itr)
		{
			m_ctrlLinearMember.SetItemText(i , 0 , itr->first.c_str());
			m_ctrlLinearMember.SetItemData(i , 0 , CEditLinearMemberShapeDlg::NONE_ITEM);

			m_ctrlLinearMember.SetItemText(i , 1 , itr->second->Shape.c_str());
			m_ctrlLinearMember.SetCellType(i, 1 , RUNTIME_CLASS(CGridCellCombo));
			CGridCellCombo* pCellCombo = (CGridCellCombo*)m_ctrlLinearMember.GetCell(i , 1);
			if(pCellCombo)
			{
				CStringArray ar;
				ar.Add(_T("ANGLE"));
				ar.Add(_T("CHANNEL"));
				ar.Add(_T("TEE"));
				ar.Add(_T("WFB"));
				
				pCellCombo->SetOptions(ar);
			}

			oss.str(_T(""));
			oss << itr->second->Height*dMultiplier;
			m_ctrlLinearMember.SetItemText(i , 2 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->Width*dMultiplier;
			m_ctrlLinearMember.SetItemText(i , 3 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->t1*dMultiplier;
			m_ctrlLinearMember.SetItemText(i , 4 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->t2*dMultiplier;
			m_ctrlLinearMember.SetItemText(i , 5 , oss.str().c_str());
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
void CEditLinearMemberShapeDlg::OnBnClickedNewLinearMember()
{
	CString sSectionName;
	GetDlgItemText(IDC_EDIT_SECTION_NAME,sSectionName);
	if(sSectionName.IsEmpty())
	{
		AfxMessageBox(_T("Section Name is empty!!!") , MB_OK | MB_ICONWARNING);
		return;
	}
	if(CSmartSteelDoc::m_oShapeParamMap.end() != CSmartSteelDoc::m_oShapeParamMap.find(sSectionName.operator LPCTSTR()))
	{
		AfxMessageBox(_T("Section Name is already existing\nPlease, input another section name!!!") , MB_OK | MB_ICONWARNING);
		return;
	}

	if(TRUE == m_ctrlLinearMember.SetRowCount( m_ctrlLinearMember.GetRowCount() + 1 ))
	{
		m_ctrlLinearMember.SetItemText(m_ctrlLinearMember.GetRowCount() - 1, 0 , sSectionName);
		m_ctrlLinearMember.SetCellType(m_ctrlLinearMember.GetRowCount() - 1, 1 , RUNTIME_CLASS(CGridCellCombo));
		CGridCellCombo* pCellCombo = (CGridCellCombo*)m_ctrlLinearMember.GetCell(m_ctrlLinearMember.GetRowCount() - 1 , 1);
		if(pCellCombo)
		{
			CStringArray ar;
			ar.Add(_T("ANGLE"));
			ar.Add(_T("CHANNEL"));
			ar.Add(_T("TEE"));
			ar.Add(_T("WFB"));

			pCellCombo->SetOptions(ar);
		}

		m_ctrlLinearMember.SetItemData(m_ctrlLinearMember.GetRowCount() - 1 , 0 , CEditLinearMemberShapeDlg::NEW_ITEM);
	}
}

/**
	@brief	delete selected linear member shape
	@author	humkyung
	@date	2013.08.12
*/
void CEditLinearMemberShapeDlg::OnBnClickedDeleteLinearMember()
{
	CAppDocData& docData = CAppDocData::GetInstance();
	
	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		CCellRange range = m_ctrlLinearMember.GetSelectedCellRange();
		for(int i = range.GetMaxRow();i >= range.GetMinRow();--i)
		{
			const CString sSectionName = m_ctrlLinearMember.GetItemText(i , 0);
			if(CEditLinearMemberShapeDlg::NEW_ITEM != m_ctrlLinearMember.GetItemData(i , 0))
			{
				/// delete item from database
				CString sSql;
				sSql.Format(_T("DELETE FROM SHAPE WHERE SectionName='%s'") , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
			}
			map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator where = CSmartSteelDoc::m_oShapeParamMap.find(sSectionName.operator LPCTSTR());
			if(where != CSmartSteelDoc::m_oShapeParamMap.end())
			{
				SAFE_DELETE(where->second);
				CSmartSteelDoc::m_oShapeParamMap.erase(where);
			}

			m_ctrlLinearMember.DeleteRow(i);
		}
		m_ctrlLinearMember.RedrawWindow();

		adodb.DBDisConnect();
	}
}

/**
	@brief	save new or modified stee shape parameter
	@author	humkyung
	@date	2013.08.13
*/
int CEditLinearMemberShapeDlg::SaveData()
{
	CAppDocData& docData = CAppDocData::GetInstance();
	const double dDivider = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1.0/1000.0) : 1.;

	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		CCellRange range = m_ctrlLinearMember.GetSelectedCellRange();
		for(int i = 1;i < m_ctrlLinearMember.GetRowCount();++i)
		{
			const CString sSectionName = m_ctrlLinearMember.GetItemText(i , 0);
			const CString sShape = m_ctrlLinearMember.GetItemText(i , 1);
			const CString sHeight = m_ctrlLinearMember.GetItemText(i , 2);
			const CString sWidth  = m_ctrlLinearMember.GetItemText(i , 3);
			const CString sT1 = m_ctrlLinearMember.GetItemText(i , 4);
			const CString sT2 = m_ctrlLinearMember.GetItemText(i , 5);
			if(sSectionName.IsEmpty() || sShape.IsEmpty() || sHeight.IsEmpty() || sWidth.IsEmpty() || sT1.IsEmpty() || sT2.IsEmpty()) continue;

			CString sSql;
			if(CEditLinearMemberShapeDlg::NEW_ITEM == m_ctrlLinearMember.GetItemData(i , 0))
			{
				/// insert item to database
				sSql.Format(_T("INSERT INTO SHAPE(SectionName,Shape,H,B,T1,T2) VALUES('%s','%s',%s,%s,%s,%s)") , sSectionName , sShape , sHeight , sWidth , sT1 , sT2);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				
				CSteelSectionBuilder::ShapeParam* pParam = new CSteelSectionBuilder::ShapeParam;
				if(NULL != pParam)
				{
					pParam->Shape = sShape.operator LPCTSTR();
					pParam->Height = ATOF_T(sHeight.operator LPCTSTR())*dDivider;
					pParam->Width  = ATOF_T(sWidth.operator LPCTSTR())*dDivider;
					pParam->t1 = ATOF_T(sT1.operator LPCTSTR())*dDivider;
					pParam->t2 = ATOF_T(sT2.operator LPCTSTR())*dDivider;
					CSmartSteelDoc::m_oShapeParamMap.insert(make_pair(sSectionName.operator LPCTSTR() , pParam));
				}
			}
			else if(CEditLinearMemberShapeDlg::MODIFIED_ITEM == m_ctrlLinearMember.GetItemData(i , 0))
			{
				/// update item
				sSql.Format(_T("UPDATE SHAPE SET Shape='%s',H=%s,B=%s,T1=%s,T2=%s WHERE SectionName='%s'") , sShape , sHeight , sWidth , sT1 , sT2 , sSectionName);
				adodb.ExecuteQuery(sSql.operator LPCTSTR());
				map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator where = CSmartSteelDoc::m_oShapeParamMap.find(sSectionName.operator LPCTSTR());
				if(where != CSmartSteelDoc::m_oShapeParamMap.end())
				{
					where->second->Shape = sShape.operator LPCTSTR();
					where->second->Height = ATOF_T(sHeight.operator LPCTSTR())*dDivider;
					where->second->Width  = ATOF_T(sWidth.operator LPCTSTR())*dDivider;
					where->second->t1 = ATOF_T(sT1.operator LPCTSTR())*dDivider;
					where->second->t2 = ATOF_T(sT2.operator LPCTSTR())*dDivider;
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
BOOL CEditLinearMemberShapeDlg::OnNotify(WPARAM wParam, LPARAM lParam, LRESULT* pResult)
{
	if(wParam == m_ctrlLinearMember.GetDlgCtrlID())
	{
		NM_GRIDVIEW* pGridView = (NM_GRIDVIEW*)(lParam);

		if(GVN_ENDLABELEDIT == pGridView->hdr.code)
		{
			CGridCellBase* pCell = m_ctrlLinearMember.GetCell(pGridView->iRow, pGridView->iColumn);
			if(pCell)
			{
				pCell->SetTextClr(RGB(255,0,0));
				pCell->EndEdit();
			}
			if(CEditLinearMemberShapeDlg::NEW_ITEM != m_ctrlLinearMember.GetItemData(pGridView->iRow , pGridView->iColumn))
			{
				m_ctrlLinearMember.SetItemData(pGridView->iRow , pGridView->iColumn , CEditLinearMemberShapeDlg::MODIFIED_ITEM);
			}
			return TRUE;
		}
	}

	return CSmartSteelPropertyPage::OnNotify(wParam, lParam, pResult);
}
