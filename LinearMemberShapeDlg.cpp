// LinearMemberShapeDlg.cpp : implementation file
//

#include "stdafx.h"

#include <IsTools.h>
#include <FileTools.h>
#include <IsPlane3d.h>
#include <ado/ADODB.h>
#include <GridCtrl/GridCellCombo.h>

#include "LinearMemberShapeDlg.h"
#include "AppDocData.h"

#include <algorithm>
using namespace std;

// CLinearMemberShapeDlg dialog

IMPLEMENT_DYNAMIC(CLinearMemberShapeDlg, CDialog)

CLinearMemberShapeDlg::CLinearMemberShapeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CLinearMemberShapeDlg::IDD, pParent)
{

}

CLinearMemberShapeDlg::~CLinearMemberShapeDlg()
{
}

void CLinearMemberShapeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CLinearMemberShapeDlg, CDialog)
	ON_BN_CLICKED(IDOK, &CLinearMemberShapeDlg::OnBnClickedOk)
END_MESSAGE_MAP()


// CLinearMemberShapeDlg message handlers

/**
	@brief	return new linear member count
	@author	humkyung
	@date	2013.08.09
*/
const int CLinearMemberShapeDlg::GetNewLinearMemberCount() const
{
	return m_oNewLinearMemberList.size();
}

/**
	@brief	add new linear member section name
	@author	humkyung
	@date	2013.08.09
*/
int CLinearMemberShapeDlg::AddNewLinearMember(const STRING_T& sSection)
{
	if(m_oNewLinearMemberList.end() == find(m_oNewLinearMemberList.begin() , m_oNewLinearMemberList.end() , sSection))
	{
		m_oNewLinearMemberList.push_back(sSection);
	}

	return ERROR_SUCCESS;
}

BOOL CLinearMemberShapeDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CAppDocData& docData = CAppDocData::GetInstance();
	const double dMultiplier = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1000.0) : 1.;

	CRect rect;

	GetDlgItem(IDC_STATIC_EXISTING_MEMBER)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	rect.top += 10;
	rect.DeflateRect(10 , 10);
	if(TRUE == m_ctrlExistingMember.Create(rect ,this , 0x101))
	{
		m_ctrlExistingMember.SetFixedRowCount(1);
		m_ctrlExistingMember.SetFixedColumnCount(1);
		m_ctrlExistingMember.SetRowCount(CSmartSteelDoc::m_oShapeParamMap.size() + 1);
		m_ctrlExistingMember.SetColumnCount(6);

		m_ctrlExistingMember.SetItemText(0 , 0 , _T("Section Name"));
		m_ctrlExistingMember.SetColumnWidth(0 , 120);
		m_ctrlExistingMember.SetItemText(0 , 1 , _T("Shape"));
		m_ctrlExistingMember.SetColumnWidth(1 , 70);
		m_ctrlExistingMember.SetItemText(0 , 2 , _T("H"));
		m_ctrlExistingMember.SetColumnWidth(2 , 70);
		m_ctrlExistingMember.SetItemText(0 , 3 , _T("B"));
		m_ctrlExistingMember.SetColumnWidth(3 , 70);
		m_ctrlExistingMember.SetItemText(0 , 4 , _T("T1"));
		m_ctrlExistingMember.SetColumnWidth(4 , 70);
		m_ctrlExistingMember.SetItemText(0 , 5 , _T("T2"));
		m_ctrlExistingMember.SetColumnWidth(5 , 70);

		m_ctrlExistingMember.SetEditable(FALSE);

		int i = 1;
		OSTRINGSTREAM_T oss;
		oss.precision( 3 );
		for(map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator itr = CSmartSteelDoc::m_oShapeParamMap.begin();itr != CSmartSteelDoc::m_oShapeParamMap.end();++itr)
		{
			m_ctrlExistingMember.SetItemText(i , 0 , itr->first.c_str());

			m_ctrlExistingMember.SetItemText(i , 1 , itr->second->Shape.c_str());

			oss.str(_T(""));
			oss << itr->second->Height*dMultiplier;
			m_ctrlExistingMember.SetItemText(i , 2 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->Width*dMultiplier;
			m_ctrlExistingMember.SetItemText(i , 3 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->t1*dMultiplier;
			m_ctrlExistingMember.SetItemText(i , 4 , oss.str().c_str());
			oss.str(_T(""));
			oss << itr->second->t2*dMultiplier;
			m_ctrlExistingMember.SetItemText(i , 5 , oss.str().c_str());
			++i;
		}
	}

	GetDlgItem(IDC_STATIC_NEW_MEMBER)->GetWindowRect(&rect);
	ScreenToClient(&rect);
	rect.top += 10;
	rect.DeflateRect(10 , 10);
	if(TRUE == m_ctrlNewMember.Create(rect ,this , 0x102))
	{
		m_ctrlNewMember.SetFixedRowCount(1);
		m_ctrlNewMember.SetFixedColumnCount(1);
		m_ctrlNewMember.SetRowCount(m_oNewLinearMemberList.size() + 1);
		m_ctrlNewMember.SetColumnCount(6);

		m_ctrlNewMember.SetItemText(0 , 0 , _T("Section Name"));
		m_ctrlNewMember.SetColumnWidth(0 , 120);
		m_ctrlNewMember.SetItemText(0 , 1 , _T("Shape"));
		m_ctrlNewMember.SetColumnWidth(1 , 70);
		m_ctrlNewMember.SetItemText(0 , 2 , _T("H"));
		m_ctrlNewMember.SetColumnWidth(2 , 70);
		m_ctrlNewMember.SetItemText(0 , 3 , _T("B"));
		m_ctrlNewMember.SetColumnWidth(3 , 70);
		m_ctrlNewMember.SetItemText(0 , 4 , _T("T1"));
		m_ctrlNewMember.SetColumnWidth(4 , 70);
		m_ctrlNewMember.SetItemText(0 , 5 , _T("T2"));
		m_ctrlNewMember.SetColumnWidth(5 , 70);

		int i = 1;
		m_oNewLinearMemberList.sort();
		for(list<STRING_T>::iterator itr = m_oNewLinearMemberList.begin();itr != m_oNewLinearMemberList.end();++itr,++i)
		{
			m_ctrlNewMember.SetItemText(i , 0 , itr->c_str());
			
			m_ctrlNewMember.SetCellType(i, 1 , RUNTIME_CLASS(CGridCellCombo));
			CGridCellCombo* pCellCombo = (CGridCellCombo*)m_ctrlNewMember.GetCell(i , 1);
			if(pCellCombo)
			{
				CStringArray ar;
				ar.Add(_T("ANGLE"));
				ar.Add(_T("CHANNEL"));
				ar.Add(_T("TEE"));
				ar.Add(_T("WFB"));
				
				pCellCombo->SetOptions(ar);
			}
		}
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

/**
	@brief	save new shape parameter
	@author	humkyung
	@date	2013.08.10
*/
void CLinearMemberShapeDlg::OnBnClickedOk()
{
	CAppDocData& docData = CAppDocData::GetInstance();
	const double dDivider = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1.0/1000.0) : 1.;

	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		const int iRowCount = m_ctrlNewMember.GetRowCount();
		for(int row = 1;row < iRowCount;++row)
		{
			const CString sSectionName = m_ctrlNewMember.GetItemText(row , 0);
			const CString sShape = m_ctrlNewMember.GetItemText(row , 1);
			const CString sHeight = m_ctrlNewMember.GetItemText(row , 2);
			const CString sWidth  = m_ctrlNewMember.GetItemText(row , 3);
			const CString sT1 = m_ctrlNewMember.GetItemText(row , 4);
			const CString sT2 = m_ctrlNewMember.GetItemText(row , 5);

			if(sSectionName.IsEmpty() || sShape.IsEmpty() || sHeight.IsEmpty() || sWidth.IsEmpty() || sT1.IsEmpty() || sT2.IsEmpty()) continue;

			CString sSql;
			sSql.Format(_T("INSERT INTO SHAPE(SectionName,Shape,H,B,T1,T2) VALUES('%s','%s',%s,%s,%s,%s)") , sSectionName , sShape , sHeight , sWidth , sT1 , sT2);
			if(TRUE == adodb.ExecuteQuery(sSql.operator LPCTSTR()))
			{
				CSteelSectionBuilder::ShapeParam* pParam = new CSteelSectionBuilder::ShapeParam;
				if(NULL != pParam)
				{
					pParam->Shape.assign(sShape.operator LPCTSTR());
					pParam->Height = ATOF_T(sHeight)*dDivider;
					pParam->Width = ATOF_T(sWidth)*dDivider;
					pParam->t1 = ATOF_T(sT1)*dDivider;
					pParam->t2 = ATOF_T(sT2)*dDivider;
					CSmartSteelDoc::m_oShapeParamMap.insert(make_pair(sSectionName.operator LPCTSTR() , pParam));
				}
			}
		}
		adodb.DBDisConnect();
	}

	OnOK();
}
