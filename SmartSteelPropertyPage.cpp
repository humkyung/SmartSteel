// SmartSteelPropertyPage.cpp : implementation file
//

#include "stdafx.h"
#include "SmartSteelPropertyPage.h"

using namespace PropertyPage;
// CSmartSteelPropertyPage dialog

IMPLEMENT_DYNAMIC(CSmartSteelPropertyPage, CDialog)

CSmartSteelPropertyPage::CSmartSteelPropertyPage(UINT nIDTempalte , CWnd* pParent /*=NULL*/)
	: CDialog(nIDTempalte , pParent)
{

}

CSmartSteelPropertyPage::~CSmartSteelPropertyPage()
{
}

void CSmartSteelPropertyPage::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}


BEGIN_MESSAGE_MAP(CSmartSteelPropertyPage, CDialog)
END_MESSAGE_MAP()


// CSmartSteelPropertyPage message handlers

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @class      CSmartSteelPropertyPage
    @function   Create
    @return     BOOL
    @param      LPCTSTR pIniFilePath
    @param      UINT    nIDTemplate
    @param      CWnd*   pParentWnd
    @brief
******************************************************************************/
BOOL CSmartSteelPropertyPage::Create(LPCTSTR pIniFilePath , UINT nIDTemplate , CWnd* pParentWnd)
{
	m_sIniFilePath = pIniFilePath;

	return CDialog::Create(nIDTemplate , pParentWnd);
}

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @function   CSmartSteelPropertyPage::SaveData
    @return     int
    @param      LPCTSTR sSettingFilePath
    @brief
******************************************************************************/
int PropertyPage::CSmartSteelPropertyPage::SaveData()
{
	return ERROR_SUCCESS;
}

/******************************************************************************
    @author     humkyung
    @date       2013-08-15
    @function   CSmartSteelPropertyPage::LoadData
    @return     int
    @param      LPCTSTR sSettingFilePath
    @brief
******************************************************************************/
int PropertyPage::CSmartSteelPropertyPage::LoadData()
{
	return ERROR_SUCCESS;
}