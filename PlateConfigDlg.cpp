// PlateConfigDlg.cpp : implementation file
//

#include "stdafx.h"
#include "SmartSteel.h"
#include "PlateConfigDlg.h"
#include "AppDocData.h"
#include "EditLinearMemberShapeDlg.h"

#include <FileTools.h>
#include <Tokenizer.h>

#include <fstream>
using namespace std;
using namespace PropertyPage;

// CPlateConfigDlg dialog

IMPLEMENT_DYNAMIC(CPlateConfigDlg, CSmartSteelPropertyPage)

CPlateConfigDlg::CPlateConfigDlg(const CString sIniFilePath , CWnd* pParent /*=NULL*/)
	: CSmartSteelPropertyPage(CPlateConfigDlg::IDD, pParent)
{
	m_sIniFilePath = sIniFilePath;

	CAppDocData& docData = CAppDocData::GetInstance();

	m_sClass.Format(_T("%d") , docData.m_oPlateCfg.class_);
	m_sGrade = docData.m_oPlateCfg.grade_.c_str();
	m_bGenerateForWebTypeBrace = (true == docData.m_oPlateCfg.generate_for_web_type_brace);
	m_bGenerateEndPlateDependOnBeamLength = (true == docData.m_oPlateCfg.generate_endplate_depend_on_beam_length);	/// 2013.10.24 added by humkyung
	m_sGussetPlateDisplayColor = docData.m_oPlateCfg.gusset_plate_display_color_.c_str();
	m_sEndPlateDisplayColor = docData.m_oPlateCfg.end_plate_display_color_.c_str();
	m_sUnit = (UNIT::M == docData.m_oPlateCfg.unit_) ? _T("M") : _T("MM");	/// 2014.02.08 added by humkyung
	m_sMaxEdgeLengthToMerge.Format(_T("%d") , docData.m_oPlateCfg.max_edge_length_to_merge_);
}

CPlateConfigDlg::~CPlateConfigDlg()
{
}

void CPlateConfigDlg::DoDataExchange(CDataExchange* pDX)
{
	CSmartSteelPropertyPage::DoDataExchange(pDX);

	DDX_Text(pDX , IDC_EDIT_CLASS , m_sClass);
	DDX_Text(pDX , IDC_EDIT_GRADE , m_sGrade);
	DDX_Check(pDX , IDC_CHECK_GENERATE_FOR_WEB_TYPE_BRACE , m_bGenerateForWebTypeBrace);
	DDX_Check(pDX , IDC_CHECK_GENERATE_ENDPLATE , m_bGenerateEndPlateDependOnBeamLength);	/// 2013.10.24 added by humkyung
	DDX_Text(pDX , IDC_COMBO_GUSSET_PLATE_COLOR , m_sGussetPlateDisplayColor);
	DDX_Text(pDX , IDC_COMBO_END_PLATE_COLOR , m_sEndPlateDisplayColor);
	DDX_CBString(pDX , IDC_COMBO_UNIT , m_sUnit);	/// 2014.02.08 added by humkyung
	DDX_Text(pDX , IDC_EDIT_MAX_EDGE_LENGTH , m_sMaxEdgeLengthToMerge);	/// 2014.02.14 added by humkyung
}

BEGIN_MESSAGE_MAP(CPlateConfigDlg, CSmartSteelPropertyPage)
	ON_BN_CLICKED(IDOK, &CPlateConfigDlg::OnBnClickedOk)
	ON_BN_CLICKED(IDC_EDIT_LINEAR_MEMBER_SHAPE, &CPlateConfigDlg::OnBnClickedEditLinearMemberShape)
END_MESSAGE_MAP()


// CPlateConfigDlg message handlers

BOOL CPlateConfigDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	CAppDocData& docData = CAppDocData::GetInstance();

	const int iCount = docData.GetColorCount();
	{
		CComboBox* pComboBox = (CComboBox*)GetDlgItem(IDC_COMBO_GUSSET_PLATE_COLOR);
		for(int i = 0;i < iCount;++i)
		{
			CAppDocData::ColorQuad colorQuad = docData.GetColorAt(i);
			pComboBox->InsertString(pComboBox->GetCount() , colorQuad.name.c_str());
		}
		const int at = pComboBox->FindStringExact(0 , m_sGussetPlateDisplayColor);
		if(-1 != at) pComboBox->SetCurSel(at);
	}
	
	{
		CComboBox* pComboBox = (CComboBox*)GetDlgItem(IDC_COMBO_END_PLATE_COLOR);
		for(int i = 0;i < iCount;++i)
		{
			CAppDocData::ColorQuad colorQuad = docData.GetColorAt(i);
			pComboBox->InsertString(pComboBox->GetCount() , colorQuad.name.c_str());
		}
		const int at = pComboBox->FindStringExact(0 , m_sEndPlateDisplayColor);
		if(-1 != at) pComboBox->SetCurSel(at);
	}

	return TRUE;  // return TRUE unless you set the focus to a control
	// EXCEPTION: OCX Property Pages should return FALSE
}

void CPlateConfigDlg::OnBnClickedOk()
{
	SaveData();

	OnOK();
}

/**
	@brief	show edit linear member shape dialog
	@author	humkyung
	@date	2013.08.12
*/
void CPlateConfigDlg::OnBnClickedEditLinearMemberShape()
{
	CEditLinearMemberShapeDlg dlg;
	if(IDOK == dlg.DoModal())
	{
	}
}

/**
	@brief	save data
	@author	humkyung
	@date	2013.08.15
*/
int CPlateConfigDlg::SaveData()
{
	CAppDocData& docData = CAppDocData::GetInstance();

	UpdateData();

	{
		docData.m_oPlateCfg.class_ = ATOI_T(m_sClass);
		docData.m_oPlateCfg.grade_ = m_sGrade;
		docData.m_oPlateCfg.generate_for_web_type_brace = (TRUE == m_bGenerateForWebTypeBrace);
		docData.m_oPlateCfg.generate_endplate_depend_on_beam_length = (TRUE == m_bGenerateEndPlateDependOnBeamLength);	/// 2013.10.24 added by humkyung
		docData.m_oPlateCfg.gusset_plate_display_color_= m_sGussetPlateDisplayColor;
		docData.m_oPlateCfg.end_plate_display_color_= m_sEndPlateDisplayColor;
		docData.m_oPlateCfg.unit_ = (_T("M") == m_sUnit) ? UNIT::M : UNIT::MM;	/// 2014.02.08 added by humkyung
		docData.m_oPlateCfg.max_edge_length_to_merge_ = ATOL_T(m_sMaxEdgeLengthToMerge);	/// 2014.02.14 added by humkyung
	}

	CString str;
	str.Format(_T("%d") , docData.m_oPlateCfg.class_);
	WritePrivateProfileString(_T("Plate") , _T("class") , str , m_sIniFilePath);
	WritePrivateProfileString(_T("Plate") , _T("grade") , docData.m_oPlateCfg.grade_.c_str() , m_sIniFilePath);
	WritePrivateProfileString(_T("Generate") , _T("generate_for_web_type_brace") , (true == docData.m_oPlateCfg.generate_for_web_type_brace) ? _T("Yes") : _T("No") , m_sIniFilePath);
	WritePrivateProfileString(_T("Generate") , _T("generate_endplate_depend_on_beam_length") , (true == docData.m_oPlateCfg.generate_endplate_depend_on_beam_length) ? _T("Yes") : _T("No") , m_sIniFilePath);
	WritePrivateProfileString(_T("Plate") , _T("gusset_plate_display_color") , docData.m_oPlateCfg.gusset_plate_display_color_.c_str() , m_sIniFilePath);
	WritePrivateProfileString(_T("Plate") , _T("end_plate_display_color") , docData.m_oPlateCfg.end_plate_display_color_.c_str() , m_sIniFilePath);
	WritePrivateProfileString(_T("Database") , _T("Unit") , m_sUnit , m_sIniFilePath);	/// 2014.02.08 added by humkyung
	WritePrivateProfileString(_T("Generate") , _T("max_edge_length_to_merge") , m_sMaxEdgeLengthToMerge , m_sIniFilePath);	/// 2014.02.08 added by humkyung

	return ERROR_SUCCESS;
}