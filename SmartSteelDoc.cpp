// This MFC Samples source code demonstrates using MFC Microsoft Office Fluent User Interface 
// (the "Fluent UI") and is provided only as referential material to supplement the 
// Microsoft Foundation Classes Reference and related electronic documentation 
// included with the MFC C++ library software.  
// License terms to copy, use or distribute the Fluent UI are available separately.  
// To learn more about our Fluent UI licensing program, please visit 
// http://msdn.microsoft.com/officeui.
//
// Copyright (C) Microsoft Corporation
// All rights reserved.

// SmartSteelDoc.cpp : implementation of the CSmartSteelDoc class
//

#include "stdafx.h"
#include <assert.h>
#include "SmartSteel.h"
#include "SmartSteelView.h"
#include "MainFrm.h"
#include "SDNFAttribute.h"

//#include <tbb/parallel_for_each.h>

#include <IsTools.h>
#include <FileTools.h>
#include <IsPlane3d.h>
#include <ado/ADODB.h>
#include <util/SplitPath.h>

#include <OCCEntFactory.h>
#include <ComplexShapeEntity.h>
#include <SphereEntity.h>

#ifdef	SMART_STEEL
#include <ImportExport/ImportExport.h>
#endif

#include "AppDocData.h"
#include "SmartSteelDoc.h"
#include "OCCShapeBuilder.h"
#include "SteelConnPoint.h"
#include "SteelSectionBuilder.h"

#include "GussetPlateGenerator.h"
#include "EndPlateGenerator.h"

#include "SmartSteelPropertySheet.h"
#include "LinearMemberShapeDlg.h"
#include "WorkStatusDlg.h"

#include "PlateFile.h"

#ifdef _DEBUG
//#ifdef _DEBUG
	#undef	DBG_NEW      
	#define DBG_NEW new ( _NORMAL_BLOCK , __FILE__ , __LINE__ )
//		#define new DBG_NEW   
//	#endif
//#endif  // _DEBUG
#define new DEBUG_NEW
#endif

map<STRING_T , CGussetPlate::Param* > CSmartSteelDoc::m_oGussetPlateParamMap;
map<STRING_T , CEndPlate::Param* > CSmartSteelDoc::m_oEndPlateParamMap;
map<STRING_T , CSteelSectionBuilder::ShapeParam* > CSmartSteelDoc::m_oShapeParamMap;

// CSmartSteelDoc

IMPLEMENT_DYNCREATE(CSmartSteelDoc, OCC_3dBaseDoc)

BEGIN_MESSAGE_MAP(CSmartSteelDoc, OCC_3dBaseDoc)
	ON_COMMAND(ID_FILE_OPEN, &CSmartSteelDoc::OnFileOpen)
	ON_UPDATE_COMMAND_UI(ID_FILE_SAVE, &CSmartSteelDoc::OnUpdateFileSave)
	ON_COMMAND(ID_FILE_SAVE, &CSmartSteelDoc::OnFileSave)
	ON_COMMAND(IDS_APP_GENERATE, &CSmartSteelDoc::OnAppGenerate)
	ON_UPDATE_COMMAND_UI(IDS_APP_GENERATE, &CSmartSteelDoc::OnUpdateAppGenerate)
#ifdef	SMART_STEEL
	ON_COMMAND(IDS_EXPORT_BREP, &CSmartSteelDoc::OnExportBREP)
#endif
	ON_COMMAND(IDS_APP_EXPORT, &CSmartSteelDoc::OnAppExport)
	ON_UPDATE_COMMAND_UI(IDS_APP_EXPORT, &CSmartSteelDoc::OnUpdateAppExport)
	ON_COMMAND(IDS_APP_OPTIONS, &CSmartSteelDoc::OnAppOptions)
	ON_UPDATE_COMMAND_UI(IDS_OCC_VIEW_SHADED, &CSmartSteelDoc::OnUpdateOCCViewShaded)
	ON_COMMAND(IDS_OCC_VIEW_SHADED, &CSmartSteelDoc::OnOCCViewShaded)
	ON_UPDATE_COMMAND_UI(IDS_OCC_VIEW_WIREFRAME, &CSmartSteelDoc::OnUpdateOCCViewWireFrame)
	ON_COMMAND(IDS_OCC_VIEW_WIREFRAME, &CSmartSteelDoc::OnOCCViewWireframe)
END_MESSAGE_MAP()


// CSmartSteelDoc construction/destruction

CSmartSteelDoc::CSmartSteelDoc() : m_iAppOperation(CSmartSteelDoc::NONE_OP) , m_eDisplayMode(AIS_Shaded)
{
	CAppDocData& docData = CAppDocData::GetInstance();

	docData.m_oPlateCfg.class_ = 3;
	docData.m_oPlateCfg.grade_ = _T("AC30");
	docData.m_oPlateCfg.gusset_plate_display_color_ = _T("magenta");
	docData.m_oPlateCfg.end_plate_display_color_ = _T("gray80");
	docData.m_oPlateCfg.generate_for_web_type_brace = false;
	docData.m_oPlateCfg.generate_endplate_depend_on_beam_length = false;	/// 2013.10.24 added by humkyung
	docData.m_oPlateCfg.unit_ = UNIT::M;	/// 2014.02.08 added by humkyung
	docData.m_oPlateCfg.max_edge_length_to_merge_ = 600;	/// 2014.02.14 added by humkyung
}

CSmartSteelDoc::~CSmartSteelDoc()
{
}

/**
	@brief
	@author	humkyung
*/
BOOL CSmartSteelDoc::OnNewDocument()
{
	if (!OCC_3dBaseDoc::OnNewDocument())
		return FALSE;

	CAppDocData& docData = CAppDocData::GetInstance();
	docData.LoadColorList();

	TCHAR szBuf[MAX_PATH + 1]={'\0',};
	m_sIniFilePath = CFileTools::GetCommonAppDataPath();
	if(_T('\\') != m_sIniFilePath[m_sIniFilePath.GetLength() - 1]) m_sIniFilePath += _T("\\");
	if(!CFileTools::ExistFolder(m_sIniFilePath + PRODUCT_PUBLISHER + CString(_T("\\")) + PRODUCT_NAME))
	{
		CFileTools::CreateFolder(m_sIniFilePath + PRODUCT_PUBLISHER + CString(_T("\\")) + PRODUCT_NAME);
	}
	m_sIniFilePath += PRODUCT_PUBLISHER + CString(_T("\\")) + PRODUCT_NAME + CString(_T("\\")) + CString(docData.GetProjectName().c_str()) + _T("\\") + PRODUCT_NAME + CString(_T(".ini"));
	
	docData.m_oPlateCfg.class_ = GetPrivateProfileInt(_T("Plate") , _T("class") , 3 , m_sIniFilePath);
	if(GetPrivateProfileString(_T("Plate") , _T("grade") , _T("") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.grade_.assign(szBuf);
	}
	else
	{
		docData.m_oPlateCfg.grade_ = _T("");
	}

	docData.m_oPlateCfg.generate_for_web_type_brace = true;
	if(GetPrivateProfileString(_T("Generate") , _T("generate_for_web_type_brace") , _T("") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.generate_for_web_type_brace = (0 == STRICMP_T(szBuf , _T("Yes")));
	}

	/// read option of generate endplate depend on beam length - 2013.10.24 added by humkyung
	docData.m_oPlateCfg.generate_endplate_depend_on_beam_length = false;
	if(GetPrivateProfileString(_T("Generate") , _T("generate_endplate_depend_on_beam_length") , _T("No") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.generate_endplate_depend_on_beam_length = (0 == STRICMP_T(szBuf , _T("Yes")));
	}
	/// up to here

	if(GetPrivateProfileString(_T("Plate") , _T("gusset_plate_display_color") , _T("") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.gusset_plate_display_color_.assign(szBuf);
	}
	
	if(GetPrivateProfileString(_T("Plate") , _T("end_plate_display_color") , _T("") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.end_plate_display_color_.assign(szBuf);
	}

	/// get database unit from setting file - 2014.02.08 added by humkyung
	if(GetPrivateProfileString(_T("Database") , _T("Unit") , _T("M") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.unit_ = (_T("M") == CString(szBuf)) ? UNIT::M : UNIT::MM;
	}
	/// up to here

	/// get max edge length to merge from setting file - 2014.02.14 added by humkyung
	if(GetPrivateProfileString(_T("Generate") , _T("max_edge_length_to_merge") , _T("600") , szBuf , MAX_PATH , m_sIniFilePath))
	{
		docData.m_oPlateCfg.max_edge_length_to_merge_ = ATOL_T(szBuf);
	}
	/// up to here
	
	return TRUE;
}

/**
	@brief	delete doc items
	@author	humkyung
	@date	2013.07.05
*/
void CSmartSteelDoc::DeletePlateContents()
{
	for(list<CSteelConnPoint*>::iterator itr = m_oConnPntList.begin();itr != m_oConnPntList.end();++itr)
	{
		SAFE_DELETE(*itr);
	}
	m_oConnPntList.clear();
	for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
	{
		SAFE_DELETE(*itr);
	}
	m_oGussetPlateList.clear();
	for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
	{
		SAFE_DELETE(*itr);
	}
	m_oEndPlateList.clear();
}

/**
	@brief	delete doc items
	@author	humkyung
	@date	2013.07.05
*/
void CSmartSteelDoc::DeleteContents()
{
	if(NULL != m_hAISContext)
	{
		m_hAISContext->CloseAllContexts();
		m_hAISContext->RemoveAll();
	}

	for(map<STRING_T , CGussetPlate::Param* >::iterator itr = m_oGussetPlateParamMap.begin();itr != m_oGussetPlateParamMap.end();++itr)
	{
		SAFE_DELETE(itr->second);
	}
	m_oGussetPlateParamMap.clear();

	for(map<STRING_T , CEndPlate::Param* >::iterator itr = m_oEndPlateParamMap.begin();itr != m_oEndPlateParamMap.end();++itr)
	{
		SAFE_DELETE(itr->second);
	}
	m_oEndPlateParamMap.clear();
	
	for(map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator itr = m_oShapeParamMap.begin();itr != m_oShapeParamMap.end();++itr)
	{
		SAFE_DELETE(itr->second);
	}
	m_oShapeParamMap.clear();

	for(map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator itr = m_oSDNFElmMap.begin();itr != m_oSDNFElmMap.end();++itr)
	{
		SAFE_DELETE(itr->second);
	}
	m_oSDNFElmMap.clear();

	for(tr1::unordered_map<STRING_T,OCC::CComplexShapeEntity* >::iterator itr = m_oLinearShapeMap.begin();itr != m_oLinearShapeMap.end();++itr)
	{
		SAFE_DELETE(itr->second);
	}
	m_oLinearShapeMap.clear();

	for(list<OCC::CComplexShapeEntity* >::iterator itr = m_oSteelShapeList.begin();itr != m_oSteelShapeList.end();++itr)
	{
		SAFE_DELETE(*itr);
	}
	m_oSteelShapeList.clear();

	DeletePlateContents();

	for(list<CSDNFAttribute*>::iterator itr = m_oAttrList.begin();itr != m_oAttrList.end();++itr)
	{
		SAFE_DELETE(*itr);
	}
	m_oAttrList.clear();
	
	OCC_3dBaseDoc::DeleteContents();
}

// CSmartSteelDoc serialization

void CSmartSteelDoc::Serialize(CArchive& ar)
{
	if (ar.IsStoring())
	{
		// TODO: add storing code here
	}
	else
	{
		// TODO: add loading code here
	}
}


// CSmartSteelDoc diagnostics

#ifdef _DEBUG
void CSmartSteelDoc::AssertValid() const
{
	OCC_3dBaseDoc::AssertValid();
}

void CSmartSteelDoc::Dump(CDumpContext& dc) const
{
	OCC_3dBaseDoc::Dump(dc);
}
#endif //_DEBUG

template <typename T> struct invoker 
{
  void operator()(T& it) const {it();}
};

/**
	@brief	check to need to split given gusset plate
	@author	humkyung
	@date	2014.05.21
	@return	true
*/
static bool NeedToSplitGussetPlate(CGussetPlate* pGussetPlate)
{
	assert(pGussetPlate && "pGussetPlate is NULL");
	if(pGussetPlate)
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		if(pGussetPlate->GetConnPnt()->GetMemberSize() > 2)
		{
			const int iMemberCount = pGussetPlate->GetConnPnt()->GetMemberSize();
			int i = 0;
			/// need to all member type is VBRACE except first one
			for(i = 1;i < iMemberCount;++i)
			{
				if(CSDNFLinearMember::VBRACE != pGussetPlate->GetConnPnt()->GetMemberAt(i)->Type()) break;
			}
			/// up to here

			return ((iMemberCount == i) && pGussetPlate->GetMaximumEdgeLength() > MAXIMUM_EDGE_LEGNTH);
		}
	}

	return false;
}

/**
	@brief	do post process(check maximum length of edge and end plate overlapping)
	@author	humkyung
	@date	2013.08.01
*/
int CSmartSteelDoc::DoPostProcess()
{
	CMainFrame* pFrameWnd = CMainFrame::GetInstance();
	CAppDocData& docData = CAppDocData::GetInstance();
	{
		list<CGussetPlate*> oGussetPlateList;
		CGussetPlateGenerator& generator = CGussetPlateGenerator::GetInstance();
		for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();/*++itr*/)
		{
			/// 2014.03.25 added by humkyung
			if(CSteelConnPoint::HBRACE_TO_HBRACE == (*itr)->GetConnPnt()->Type())
			{
				++itr;
				continue;
			}
			/// up to here
			if(NeedToSplitGussetPlate(*itr))
			{
				for(int i = 1;i < (*itr)->GetConnPnt()->GetMemberSize();++i)
				{
					list<CGussetPlate*> lst;
					CSteelConnPoint* pConnPnt = this->CreateConnPoint((*itr)->GetConnPnt()->Type() , (*itr)->GetConnPnt()->origin());
					pConnPnt->Add((*itr)->GetConnPnt()->GetMemberAt(0));
					pConnPnt->Add((*itr)->GetConnPnt()->GetMemberAt(i));

					generator.Generate(lst , pConnPnt , this);
					oGussetPlateList.insert(oGussetPlateList.end() , lst.begin() , lst.end());
				}

				SAFE_DELETE(*itr);
				itr = m_oGussetPlateList.erase(itr);
				
				continue;
			}
			++itr;
		}
		m_oGussetPlateList.insert(m_oGussetPlateList.end() , oGussetPlateList.begin() , oGussetPlateList.end());

		for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
		{
			if(ERROR_SUCCESS == (*itr)->Display(this->GetAISContext()))
			{
				pFrameWnd->GetFileView().Add(*itr);
			}
		}
	}

	/// end plate
	for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();/*++itr*/)
	{
		Bnd_Box box1 = (*itr)->BoundBox();

		bool bOverlap = false;
		list<CEndPlate*>::iterator jtr = itr;
		for(++jtr;jtr != m_oEndPlateList.end();++jtr)
		{
			Bnd_Box box2 = (*jtr)->BoundBox();
			if(!box1.IsOut(box2))
			{
				SAFE_DELETE(*jtr);
				m_oEndPlateList.erase(jtr);
				bOverlap = true;
				break;
			}
		}
		if(true == bOverlap)
		{
			SAFE_DELETE(*itr);
			itr = m_oEndPlateList.erase(itr);
		}
		else
		{
			++itr;
		}
	}
	for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
	{
		if(ERROR_SUCCESS == (*itr)->Display(this->GetAISContext()))
		{
			pFrameWnd->GetFileView().Add(*itr);
		}
	}

	return ERROR_SUCCESS;
}

// CSmartSteelDoc commands
UINT CSmartSteelDoc::ThreadEntry(LPVOID pParam)
{
	CMainFrame* pFrameWnd = CMainFrame::GetInstance();
	OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
	CWorkStatusDlg* pProgressBar = CWorkStatusDlg::GetInstance();

	if(FILE_OPEN == m_iAppOperation)
	{
		const int iCount = m_sdnf.GetElementCount();
		for(int i = 0;i < iCount;++i)
		{
			CSDNFElement* pElm = m_sdnf.GetElementAt(i);
			if((NULL != pElm) && (pElm->IsKindOf(CSDNFLinearMemberPacket::TypeString())))
			{
				CSDNFLinearMemberPacket* pPacket = static_cast<CSDNFLinearMemberPacket*>(pElm);
				pProgressBar->UpdateWorkStatus(_T("Reading...") , 0);
				for(int j = 0;j < pPacket->GetElementCount();++j)
				{
					CSDNFLinearMember* pLinearMember = static_cast<CSDNFLinearMember*>(pPacket->GetElementAt(j));
					const CSDNFLinearMember::ElmType type = pLinearMember->Type();
					/*if( (_T("0001008134") != pLinearMember->MemberID()) && (_T("0001008138") != pLinearMember->MemberID()) && 
						(_T("0001007243") != pLinearMember->MemberID()) && (_T("0001008135") != pLinearMember->MemberID()) && 
						(_T("0001008141") != pLinearMember->MemberID()) && (_T("0001008140") != pLinearMember->MemberID()) &&
						(_T("0001008142") != pLinearMember->MemberID()) && (_T("0001007283") != pLinearMember->MemberID()) &&
						(_T("0001007343") != pLinearMember->MemberID()) && (_T("0001006213") != pLinearMember->MemberID()) &&
						(_T("0001007260") != pLinearMember->MemberID()) && (_T("0001007621") != pLinearMember->MemberID()) &&
						(_T("0001007622") != pLinearMember->MemberID()) && (_T("0001008144") != pLinearMember->MemberID())) continue;*/
					/// store linear member - 2013.06.19 added by humkyung
					map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator where = m_oSDNFElmMap.find(pLinearMember->Type());
					if(where == m_oSDNFElmMap.end())
					{
						list<CSDNFLinearMember*>* pList = new list<CSDNFLinearMember*>;
						pList->push_back( pLinearMember );
						m_oSDNFElmMap.insert(make_pair(pLinearMember->Type() , pList));
					}
					else
					{
						where->second->push_back(pLinearMember);
					}
					/// up to here

					COCCShapeBuilder oBuilder;
					TopoDS_Shape aShape = oBuilder.Shape(CIsLine3d(pLinearMember->start(CSDNFElement::METER) , pLinearMember->end(CSDNFElement::METER)));
					try
					{
						if(!aShape.IsNull())
						{
							OCC::CComplexShapeEntity* pComplexShape = (OCC::CComplexShapeEntity*)factory.GetEntity( OCC::CComplexShapeEntity::TypeString() );
							if(pComplexShape)
							{
								pComplexShape->SetColor(GetColorStringOf(type).c_str());
								pComplexShape->m_hShape = aShape;

								STRING_T sType = pLinearMember->ElmTypeString();
								STRING_T sStart , sEnd;
								{
									OSTRINGSTREAM_T oss;
									oss.precision( 3 );									/// 소수점 자릿수 3
									oss.setf(ios_base:: fixed, ios_base:: floatfield);	/// 소수점 방식으로 표시됨
									oss << pLinearMember->start().x() << _T(",") << pLinearMember->start().y() << _T(",") << pLinearMember->start().z();
									sStart = oss.str();

									oss.str(_T(""));
									oss << pLinearMember->end().x() << _T(",") << pLinearMember->end().y() << _T(",") << pLinearMember->end().z();
									sEnd = oss.str();
								}
								CSDNFAttribute* pAttr = new CSDNFAttribute(sType , pLinearMember->section() , pLinearMember->grade() ,  sStart , sEnd);
								{
									pAttr->id() = pLinearMember->MemberID();
									pAttr->CardinalPnt() = pLinearMember->CardinalPoint();
									pAttr->Rotation() = pLinearMember->rotation();
									pComplexShape->AddAttribute( pAttr );
									m_oAttrList.push_back(pAttr);
								}
								m_oLinearShapeMap.insert(make_pair(pLinearMember->MemberID() , pComplexShape));

								pFrameWnd->GetFileView().Add(pLinearMember);
							}
						}
					}
					catch(...)
					{
					}

					try
					{
						CSteelSectionBuilder oSteelShapeBuilder;
						oSteelShapeBuilder.Build(pLinearMember , this);
						if(!oSteelShapeBuilder.m_oSectionPntList.empty())
						{
							CIsPoint3d start = pLinearMember->start(CSDNFElement::METER);
							CIsVect3d dir = pLinearMember->end(CSDNFElement::METER) - start;
							const double thickness = dir.Length();
							dir.Normalize();

							for(vector<CIsPoint3d>::iterator itr = oSteelShapeBuilder.m_oSectionPntList.begin();itr != oSteelShapeBuilder.m_oSectionPntList.end();++itr)
							{
								(*itr) += start;
							}
							COCCShapeBuilder oBuilder;
							TopoDS_Shape aShape = oBuilder.Shape(oSteelShapeBuilder.m_oSectionPntList , dir , thickness);
							if(!aShape.IsNull())
							{
								OCC::CComplexShapeEntity* pComplexShape = (OCC::CComplexShapeEntity*)factory.GetEntity( OCC::CComplexShapeEntity::TypeString() );
								if(pComplexShape)
								{
									pComplexShape->SetColor(GetColorStringOf(type).c_str());
									pComplexShape->m_hShape = aShape;

									STRING_T sType = pLinearMember->ElmTypeString();
									STRING_T sStart , sEnd;
									{
										OSTRINGSTREAM_T oss;
										oss.precision( 3 );									/// 소수점 자릿수 3
										oss.setf(ios_base:: fixed, ios_base:: floatfield);	/// 소수점 방식으로 표시됨
										oss << pLinearMember->start().x() << _T(",") << pLinearMember->start().y() << _T(",") << pLinearMember->start(CSDNFElement::METER).z();
										sStart = oss.str();

										oss.str(_T(""));
										oss << pLinearMember->end().x() << _T(",") << pLinearMember->end().y() << _T(",") << pLinearMember->end(CSDNFElement::METER).z();
										sEnd = oss.str();
									}
									CSDNFAttribute* pAttr = new CSDNFAttribute(sType , pLinearMember->section() , pLinearMember->grade() ,  sStart , sEnd);
									{
										pAttr->id() = pLinearMember->MemberID();
										pAttr->CardinalPnt() = pLinearMember->CardinalPoint();
										pAttr->Rotation() = pLinearMember->rotation();
										pComplexShape->AddAttribute( pAttr );
										m_oAttrList.push_back(pAttr);
									}
									m_oSteelShapeList.push_back(pComplexShape);
									pComplexShape->SetTransparency(0.0);
								}
							}
							else
							{
								CString str;
								str.Format(_T("Fail to build shape about %s") , pLinearMember->MemberID().c_str());
								pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(str.operator LPCTSTR()) , MessageType::MESSAGE_WARNING);
							}
						}
					}
					catch(const exception& ex)
					{
						AfxMessageBox(ex.what() , MB_OK | MB_ICONERROR);
					}

					pProgressBar->UpdateWorkStatus(_T("Reading...") , int((double(j) / double(pPacket->GetElementCount()))*100));
				}
				pProgressBar->UpdateWorkStatus(_T("Reading...") , int(100));

				/// display linear member
				pProgressBar->UpdateWorkStatus(_T("Displaying...") , int(0));
				const int iTotal = m_oLinearShapeMap.size() + m_oSteelShapeList.size();
				int iIndex = 0;
				for(tr1::unordered_map<STRING_T,OCC::CComplexShapeEntity* >::iterator itr = m_oLinearShapeMap.begin();itr != m_oLinearShapeMap.end();++itr)
				{
					(itr->second)->Display(this->GetAISContext() , AIS_Shaded);
					pProgressBar->UpdateWorkStatus(_T("Displaying...") , int((double(iIndex++) / double(iTotal)*100)));
				}
				for(list<OCC::CComplexShapeEntity* >::iterator itr = m_oSteelShapeList.begin();itr != m_oSteelShapeList.end();++itr)
				{
					(*itr)->Display(this->GetAISContext() , AIS_Shaded);
					pProgressBar->UpdateWorkStatus(_T("Displaying...") , int((double(iIndex++) / double(iTotal)*100)));
				}
				pProgressBar->UpdateWorkStatus(_T("Displaying...") , int(100));
				/// up to here

				int iColumnCount = 0 , iBeamCount = 0 , iHBraceCount = 0 , iVBraceCount = 0;
				for(map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator itr = m_oSDNFElmMap.begin();itr != m_oSDNFElmMap.end();++itr)
				{
					if(CSDNFLinearMember::COLUMN == itr->first)
					{
						iColumnCount = itr->second->size();
					}
					else if(CSDNFLinearMember::BEAM == itr->first)
					{
						iBeamCount = itr->second->size();
					}
					else if(CSDNFLinearMember::HBRACE == itr->first)
					{
						iHBraceCount = itr->second->size();
					}
					else if(CSDNFLinearMember::VBRACE == itr->first)
					{
						iVBraceCount = itr->second->size();
					}
				}
				CString str;
				str.Format(_T("Columns(%d),Beams(%d),Hor. Braces(%d),Ver. Braces(%d) are loaded") , iColumnCount , iBeamCount , iHBraceCount , iVBraceCount);
				pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(str.operator LPCTSTR()) , MessageType::MESSAGE_INFO);

				/// load plate file - 2014.09.03 added by humkyung
				CSplitPath path(this->GetTitle());
				const CString sPlateFilePath(path.GetDrive() + path.GetDirectory() + path.GetFileName() + _T(".mdt"));
				this->LoadPlate(sPlateFilePath);
				for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
				{
					if(ERROR_SUCCESS == (*itr)->Display(this->GetAISContext()))
					{
						pFrameWnd->GetFileView().Add(*itr);
					}
					else
					{
						pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(_T("Fail to display gusset plate")) , MessageType::MESSAGE_WARNING);
					}
				}
				/// up to here
			}
		}
	}
	else	/// generate gusset/end plate
	{
		int i = 0 , iCount = int(m_oConnPntList.size());
		pProgressBar->UpdateWorkStatus(_T("Generating...") , int(0));
		pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(_T("Generating gusset/end plates...")) , MessageType::MESSAGE_INFO);

		/// generate gusset plate
		CGussetPlateGenerator& oGussetGenerator = CGussetPlateGenerator::GetInstance();
		CEndPlateGenerator& oEndGenerator = CEndPlateGenerator::GetInstance();
		for(list<CSteelConnPoint*>::iterator itr = m_oConnPntList.begin();itr != m_oConnPntList.end();++itr)
		{
			if(CSteelConnPoint::COLUMN_TO_BEAM == (*itr)->Type())
			{
				list<CEndPlate*> oEndPlateList;
				oEndGenerator.Generate(oEndPlateList , (*itr) , this);
				m_oEndPlateList.insert(m_oEndPlateList.end() , oEndPlateList.begin() , oEndPlateList.end());
			}
			else
			{
				list<CGussetPlate*> oGussetPlateList;
				oGussetGenerator.Generate(oGussetPlateList , (*itr) , this);
				m_oGussetPlateList.insert(m_oGussetPlateList.end() , oGussetPlateList.begin() , oGussetPlateList.end());
			}
			
			(*itr)->Display(this->GetAISContext());

			pProgressBar->UpdateWorkStatus(_T("Generating...") , int((double(i++) / double(iCount))*100));
		}
		
		DoPostProcess();
		
		/// up to here
		pProgressBar->UpdateWorkStatus(_T("Generating...") , int(100));

		CString str;
		str.Format(_T("Gusset Plates(%d),End Plates(%d) are generated") , m_oGussetPlateList.size() , m_oEndPlateList.size());
		pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(str.operator LPCTSTR()) , MessageType::MESSAGE_INFO);

		this->UpdateAllViews(false);
	}
	m_iAppOperation = CSmartSteelDoc::NONE_OP;

	return 0;
}

UINT CSmartSteelDoc::StatusThreadEntry(LPVOID pParam)
{
	CSmartSteelDoc* pDoc = (CSmartSteelDoc*)(pParam);
	
	CWorkStatusDlg* pProgressBar = CWorkStatusDlg::GetInstance();
	pDoc->ThreadEntry(pParam);
	InterlockedExchange((LONG*)(&(pProgressBar->m_bThreadRunning)) , FALSE);
	pProgressBar->PostMessage(WM_COMMAND , IDOK);

	return 0;
}

/**
	@brief	load gusset/end plates
	@author	humkyung
	@date	2014.09.03
	@return	int
*/
int CSmartSteelDoc::LoadPlate(const CString& sFilePath)
{
	list<CGussetPlate*> oGussetPlateList;
	list<CEndPlate*> oEndPlateList;

	CPlateFile file;
	if(ERROR_SUCCESS == file.Load(oGussetPlateList , oEndPlateList , sFilePath , this))
	{
		this->m_oGussetPlateList.insert(this->m_oGussetPlateList.begin() , oGussetPlateList.begin() , oGussetPlateList.end());
		this->m_oEndPlateList.insert(this->m_oEndPlateList.begin() , oEndPlateList.begin() , oEndPlateList.end());
	}

	return ERROR_SUCCESS;
}

/**
	@brief	check if there is missing linear member shape for SDNF file
	@author	humkyung
	@date	2013.08.09
*/
int CSmartSteelDoc::CheckLinearMember()
{
	CLinearMemberShapeDlg dlg;

	const int iCount = m_sdnf.GetElementCount();
	for(int i = 0;i < iCount;++i)
	{
		CSDNFElement* pElm = m_sdnf.GetElementAt(i);
		if((NULL != pElm) && (pElm->IsKindOf(CSDNFLinearMemberPacket::TypeString())))
		{
			CSDNFLinearMemberPacket* pPacket = static_cast<CSDNFLinearMemberPacket*>(pElm);
			for(int j = 0;j < pPacket->GetElementCount();++j)
			{
				CSDNFLinearMember* pLinearMember = static_cast<CSDNFLinearMember*>(pPacket->GetElementAt(j));
				const STRING_T sSectionName = pLinearMember->section();
				STRING_T tmp(sSectionName);
				if(tmp.empty()) continue;	/// check section name - 2013.10.25 added by humkyung
				if('"' == tmp[0]) tmp = tmp.substr(1);
				if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);

				map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator where = m_oShapeParamMap.find(tmp);
				if(where == m_oShapeParamMap.end())
				{
					dlg.AddNewLinearMember(tmp);
				}
			}
		}
	}
	
	if(dlg.GetNewLinearMemberCount() > 0)
	{
		dlg.DoModal();
	}

	return ERROR_SUCCESS;
}

/**
	@brief	read sdnf file and then show it
	@author	humkyung
*/
void CSmartSteelDoc::OnFileOpen()
{
	TCHAR szFilter[] = _T("SDNF (*.DAT)|*.DAT|All Files(*.*)|*.*||");
	CFileDialog dlg(TRUE, NULL, NULL, OFN_HIDEREADONLY, szFilter);
	if(IDOK == dlg.DoModal()) 
	{
		CWnd* pWndMain = AfxGetMainWnd();
		ASSERT(pWndMain);
		ASSERT(pWndMain->IsKindOf(RUNTIME_CLASS(CFrameWndEx)) && !pWndMain->IsKindOf(RUNTIME_CLASS(CMDIFrameWndEx)));
		CMainFrame* pFrameWnd = (CMainFrame*)pWndMain;

		DeleteContents();

		const CString strPathName = dlg.GetPathName();
		OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
		this->LoadSettingData();
#ifdef	SMART_STEEL
			GetAISContext()->DefaultDrawer ()->SetFaceBoundaryDraw(true);
			GetAISContext()->DefaultDrawer ()->FaceBoundaryAspect()->SetColor(Quantity_NameOfColor::Quantity_NOC_BLACK);
			GetAISContext()->DefaultDrawer ()->FaceBoundaryAspect()->SetWidth(0.5);
#endif

		pFrameWnd->SendMessage(DISPLAY_MESSAGE , 0 , 0);
		pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(_T("Reading model...")) , MessageType::MESSAGE_INFO);
		if(ERROR_SUCCESS == m_sdnf.Read(strPathName.operator LPCTSTR()))
		{
			this->SetTitle(strPathName);

			CheckLinearMember();

			CFileView& oFileView = pFrameWnd->GetFileView();
			oFileView.ResetContents();

			m_iAppOperation = FILE_OPEN;
			CWorkStatusDlg progressbar(pFrameWnd);
			progressbar.m_pThread = AfxBeginThread(StatusThreadEntry, this , THREAD_PRIORITY_NORMAL,0,CREATE_SUSPENDED);
			InterlockedExchange((LONG*)(&(progressbar.m_bThreadRunning)) , TRUE);
			if(IDOK == progressbar.DoModal()){}
			
			try
			{
				/// fit all views with ISO view
				POSITION pos = this->GetFirstViewPosition();
				while(pos)
				{
					CView* pView = this->GetNextView(pos);
					if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
					{
						CSmartSteelView* pSmartSteelView = static_cast<CSmartSteelView*>(pView);
						if(NULL != pSmartSteelView)
						{
							pSmartSteelView->OnOCCViewISO();
							pSmartSteelView->FitAll();
						}
					}
				}
			}
			catch(...){}
		}
	}
}

/**
	@brief	update file save command
	@author	humkyung
	@date	2014.09.03
*/
void CSmartSteelDoc::OnUpdateFileSave(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(!m_oGussetPlateList.empty() || !m_oEndPlateList.empty());
}

/**
	@brief	save gusset/end plate to file
	@author	humkyung
	@date	2014.09.03
*/
void CSmartSteelDoc::OnFileSave()
{
	CSplitPath path(this->GetTitle());
	const CString sOutputFilePath(path.GetDrive() + path.GetDirectory() + path.GetFileName() + _T(".mdt"));
	
	OFSTREAM_T ofile(sOutputFilePath.operator LPCTSTR());
	if(ofile.is_open())
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		const double dScale = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1000.0) : 1.;

		ofile.precision( 8 );									/// 소수점 자릿수 8
		ofile.setf(ios_base:: fixed, ios_base:: floatfield);	/// 소수점 방식으로 표시됨
		for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
		{
			(*itr)->Write(ofile , dScale);
		}
		for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
		{
			(*itr)->Write(ofile , dScale);
		}
		ofile.close();

		AfxMessageBox(_T("Save is done"));
	}
}

/******************************************************************************
    @brief		get volume of linear member
	@author     humkyung
    @date       2013-07-10
    @class      CSmartSteelDoc
    @function   GetVolumeOf
    @return     CIsVolume
******************************************************************************/
CIsVolume CSmartSteelDoc::GetVolumeOf(CSDNFLinearMember* pLinearMem)
{
	assert(pLinearMem && "pLinear Member is NULL");

	CIsVolume res;
	if(pLinearMem)
	{
		CSteelSectionBuilder oSteelShapeBuilder;
		oSteelShapeBuilder.Build(pLinearMem , this);
		const CIsPoint3d start = pLinearMem->start(CSDNFElement::METER);
		CIsVect3d dir = pLinearMem->end(CSDNFElement::METER) - start;
		const double thickness = dir.Length();
		dir.Normalize();

		for(vector<CIsPoint3d>::iterator itr = oSteelShapeBuilder.m_oSectionPntList.begin();itr != oSteelShapeBuilder.m_oSectionPntList.end();++itr)
		{
			(*itr) += start;
		}
		COCCShapeBuilder oBuilder;
		res = oBuilder.Volume(oSteelShapeBuilder.m_oSectionPntList , dir , thickness);
	}
	
	return res;
}

/**
	@brief	create a new connection point
	@author	humkyung
	@date	2014.08.15
*/
CSteelConnPoint* CSmartSteelDoc::CreateConnPoint(const CSteelConnPoint::ConnType& iType , const CIsPoint3d& at)
{
	CSteelConnPoint* res = new CSteelConnPoint(iType , at);
	m_oConnPntList.push_back(res);

	return res;
}

/******************************************************************************
    @brief		generate connection point
	@author     humkyung
    @date       2013-07-03
    @class      CSmartSteelDoc
    @function   GenerateConnPoint
    @return     void
******************************************************************************/
int CSmartSteelDoc::GenerateConnPoint()
{
	CConnPntGenerator generator;
	return generator.Generate(this);
}

/**
	@brief
	@author	humkyung
	@date	2013.08.04
*/
void CSmartSteelDoc::OnUpdateAppGenerate(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(m_iAppOperation == CSmartSteelDoc::NONE_OP);
}

/******************************************************************************
    @brief		generate gusset and end plate
	@author     humkyung
    @date       2013-06-19
    @class      CSmartSteelDoc
    @function   OnAppGenerate
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnAppGenerate()
{
	CWnd* pWndMain = AfxGetMainWnd();
	ASSERT(pWndMain);
	ASSERT(pWndMain->IsKindOf(RUNTIME_CLASS(CFrameWndEx)) && !pWndMain->IsKindOf(RUNTIME_CLASS(CMDIFrameWndEx))); // Not an SDI app.
	CMainFrame* pFrameWnd = (CMainFrame*)pWndMain;

	if(m_sdnf.GetElementCount() > 0)
	{
		pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(_T("Generating connection points...")) , MessageType::MESSAGE_INFO);
		DeletePlateContents();
		CFileView& oFileView = pFrameWnd->GetFileView();
		oFileView.ResetPlateContents();

		GenerateConnPoint();

		m_iAppOperation = CSmartSteelDoc::APP_GENERATE;
		CWorkStatusDlg progressbar(pFrameWnd);
		progressbar.m_pThread = AfxBeginThread(StatusThreadEntry, this , THREAD_PRIORITY_NORMAL,0,CREATE_SUSPENDED);
		InterlockedExchange((LONG*)(&(progressbar.m_bThreadRunning)) , TRUE);
		if(IDOK == progressbar.DoModal()){}
	}
}

/******************************************************************************
    @brief		check duplicate origin when type is less then given type
	@author     humkyung
    @date       2013-06-19
    @class      CSmartSteelDoc
    @function   CheckDuplicate
    @return     bool
******************************************************************************/
bool CSmartSteelDoc::CheckDuplicate(const int& iType , const CIsPoint3d& pt , CSDNFLinearMember* pMem1 , CSDNFLinearMember* pMem2)
{
	for(list<CSteelConnPoint*>::iterator itr = m_oConnPntList.begin();itr != m_oConnPntList.end();++itr)
	{
		if((*itr)->Type() < iType)
		{
			if((*itr)->origin() == pt)
			{
				if((NULL == pMem1) && (NULL == pMem2)) return true;
				if((NULL != pMem1) && (*itr)->HasMember(pMem1)) return true;
				if((NULL != pMem2) && (*itr)->HasMember(pMem2)) return true;
				return false;
			}
		}
	}

	return false;
}

/******************************************************************************
    @brief
	@author     humkyung
    @date       2013-05-29
    @class      CSmartSteelDoc
    @function   UpdateAllViews
    @return     int
******************************************************************************/
int CSmartSteelDoc::UpdateAllViews(const bool& bFitView)
{
	POSITION pos = this->GetFirstViewPosition();
	while(pos)
	{
		CView* pView = this->GetNextView(pos);
		if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
		{
			CSmartSteelView* _pView = static_cast<CSmartSteelView*>(pView);
			if(NULL != _pView)
			{
				if(true == bFitView) _pView->FitAll();
				_pView->Invalidate();
			}
		}
	}

	return ERROR_SUCCESS;
}

/**
	@brief	get occ entity from topo shap
	@author	humkyung
	@date	2013.06.04
*/
OCC::CComplexShapeEntity* CSmartSteelDoc::GetOCCEntityFrom(const TopoDS_Shape& selectedShape)
{
	for(list<OCC::CComplexShapeEntity* >::iterator itr = m_oSteelShapeList.begin();itr != m_oSteelShapeList.end();++itr)
	{
		if((*itr)->m_hShape == selectedShape)
		{
			return (*itr);
		}
	}

	for(tr1::unordered_map<STRING_T,OCC::CComplexShapeEntity* >::iterator itr = m_oLinearShapeMap.begin();itr != m_oLinearShapeMap.end();++itr)
	{
		if((itr->second)->m_hShape == selectedShape)
		{
			return (itr->second);
		}
	}

	return NULL;
}

/**
	@brief	get plate from topo shap
	@author	humkyung
	@date	2013.07.30
*/
CSteelPlate* CSmartSteelDoc::GetPlateFrom(const TopoDS_Shape& selectedShape)
{
	for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
	{
		if((*itr)->HasShape(selectedShape)) return (*itr);
	}

	for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
	{
		if((*itr)->HasShape(selectedShape)) return (*itr);
	}
	
	return NULL;
}

/**
	@brief	remove occ entity assciated to given shape

	@author	humkyung

	@date	2013.07.22
*/
int CSmartSteelDoc::RemoveOCCEntity(const TopoDS_Shape& selectedShape)
{
	for(list<CGussetPlate* >::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
	{
		if((*itr)->HasShape(selectedShape))
		{
			SAFE_DELETE(*itr);
			itr = m_oGussetPlateList.erase(itr);
			return ERROR_SUCCESS;
		}
	}

	for(list<CEndPlate* >::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
	{
		if((*itr)->HasShape(selectedShape))
		{
			SAFE_DELETE(*itr);
			itr = m_oEndPlateList.erase(itr);
			return ERROR_SUCCESS;
		}
	}

	return ERROR_SUCCESS;
}

/**
	@brief	load setting data
	@author	humkyung
	@date	2013.07.05
*/
int CSmartSteelDoc::LoadSettingData()
{
	{
		for(map<STRING_T , CGussetPlate::Param* >::iterator itr = m_oGussetPlateParamMap.begin();itr != m_oGussetPlateParamMap.end();++itr)
		{
			SAFE_DELETE(itr->second);
		}
		m_oGussetPlateParamMap.clear();

		for(map<STRING_T , CEndPlate::Param* >::iterator itr = m_oEndPlateParamMap.begin();itr != m_oEndPlateParamMap.end();++itr)
		{
			SAFE_DELETE(itr->second);
		}
		m_oEndPlateParamMap.clear();
		
		for(map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator itr = m_oShapeParamMap.begin();itr != m_oShapeParamMap.end();++itr)
		{
			SAFE_DELETE(itr->second);
		}
		m_oShapeParamMap.clear();
	}

	CAppDocData& docData = CAppDocData::GetInstance();

	CString sConnString = CString(_T("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=")) + docData.GetConfigFilePath();
	CADODB adodb;
	if(TRUE == adodb.DBConnect(sConnString.operator LPCTSTR()))
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		const double dDivider = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1.0/1000.0) : 1.;

		m_oGussetPlateParamMap.clear();

		adodb.OpenQuery(_T("SELECT * FROM GUSSET_PLATE"));
		LONG lRecordCount = 0L;
		adodb.GetRecordCount(&lRecordCount);
		for(int record = 0;record < lRecordCount;++record)
		{
			STRING_T sSectionSize;
			adodb.GetFieldValue(record , _T("SectionName") , &sSectionSize);
			sSectionSize = CString(sSectionSize.c_str()).MakeUpper().operator LPCTSTR();
			map<STRING_T , CGussetPlate::Param* >::iterator where = m_oGussetPlateParamMap.find(sSectionSize);
			if(where == m_oGussetPlateParamMap.end())
			{
				CGussetPlate::Param* pValueMap = new CGussetPlate::Param;
				if(NULL != pValueMap)
				{
					STRING_T sValue;

					adodb.GetFieldValue(record , _T("A") , &sValue);
					pValueMap->A = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("B") , &sValue);
					pValueMap->B = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("E") , &sValue);
					pValueMap->E = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("N") , &sValue);
					pValueMap->N = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("P") , &sValue);
					pValueMap->P = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("T") , &sValue);
					pValueMap->T = ATOF_T(sValue.c_str())*dDivider;

					m_oGussetPlateParamMap.insert(make_pair(sSectionSize , pValueMap));
				}
			}
		}

		m_oEndPlateParamMap.clear();
		adodb.OpenQuery(_T("SELECT * FROM END_PLATE"));
		lRecordCount = 0L;
		adodb.GetRecordCount(&lRecordCount);
		for(int record = 0;record < lRecordCount;++record)
		{
			STRING_T sSectionName;
			adodb.GetFieldValue(record , _T("SectionName") , &sSectionName);
			map<STRING_T , CEndPlate::Param* >::iterator where = m_oEndPlateParamMap.find(sSectionName);
			if(where == m_oEndPlateParamMap.end())
			{
				CEndPlate::Param* pValueMap = new CEndPlate::Param;
				if(NULL != pValueMap)
				{
					STRING_T sValue;

					pValueMap->SectionName = sSectionName;
					adodb.GetFieldValue(record , _T("K") , &sValue);
					pValueMap->K = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("M") , &sValue);
					pValueMap->M = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("T") , &sValue);
					pValueMap->T = ATOF_T(sValue.c_str())*dDivider;

					m_oEndPlateParamMap.insert(make_pair(sSectionName , pValueMap));
				}
			}
		}

		m_oShapeParamMap.clear();
		adodb.OpenQuery(_T("SELECT * FROM SHAPE"));
		lRecordCount = 0L;
		adodb.GetRecordCount(&lRecordCount);
		for(int record = 0;record < lRecordCount;++record)
		{
			STRING_T sSectionName;
			adodb.GetFieldValue(record , _T("SectionName") , &sSectionName);
			map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator where = m_oShapeParamMap.find(sSectionName);
			if(where == m_oShapeParamMap.end())
			{
				CSteelSectionBuilder::ShapeParam* pShapeParam = new CSteelSectionBuilder::ShapeParam;
				if(NULL != pShapeParam)
				{
					STRING_T sValue;

					adodb.GetFieldValue(record , _T("Shape") , &(pShapeParam->Shape));
					adodb.GetFieldValue(record , _T("H") , &sValue);
					pShapeParam->Height = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("B") , &sValue);
					pShapeParam->Width = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("T1") , &sValue);
					pShapeParam->t1 = ATOF_T(sValue.c_str())*dDivider;
					adodb.GetFieldValue(record , _T("T2") , &sValue);
					pShapeParam->t2 = ATOF_T(sValue.c_str())*dDivider;

					m_oShapeParamMap.insert(make_pair(sSectionName , pShapeParam));
				}
			}
		}
	}
	else
	{
		AfxMessageBox(_T("Can't open database") , MB_OK|MB_ICONEXCLAMATION);
		return ERROR_BAD_ENVIRONMENT;
	}

	return ERROR_SUCCESS;
}

/**
	@brief	

	@author	humkyung
*/
BOOL CSmartSteelDoc::OnOpenDocument(LPCTSTR lpszPathName)
{
	if (!OCC_3dBaseDoc::OnOpenDocument(lpszPathName))
		return FALSE;

	

	return TRUE;
}

/**
	@brief	get steel value corressponding to type and section

	@author	humkyung

	@date	2013.06.20
*/
CGussetPlate::Param* CSmartSteelDoc::GetSteelValue(const CSDNFLinearMember::ElmType& type , const STRING_T& section/* , const STRING_T& sKey*/)
{
	CGussetPlate::Param* res = NULL;
	if(CSDNFLinearMember::VBRACE == type)
	{
		STRING_T tmp(section);
		if('"' == tmp[0]) tmp = tmp.substr(1);
		if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);
		tmp = CString(tmp.c_str()).MakeUpper().operator LPCTSTR();

		map<STRING_T , CGussetPlate::Param* >::iterator where = m_oGussetPlateParamMap.find(tmp);
		if(where != m_oGussetPlateParamMap.end())
		{
			return where->second;
		}
	}
	
	OSTRINGSTREAM_T oss;
	oss << _T("Fail to get steel value about section(") << section << _T(")");
	CMainFrame* pFrameWnd = CMainFrame::GetInstance();
	pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(oss.str().c_str()) , MessageType::MESSAGE_WARNING);

	return res;
}

/**
	@brief	get steel joint length corressponding to type and section

	@author	humkyung

	@date	2013.06.20
*/
double CSmartSteelDoc::GetSteelJointLength(const CSDNFLinearMember::ElmType& type , const STRING_T& section)
{
	double res = -1.0;
	if(CSDNFLinearMember::VBRACE == type)
	{
		CGussetPlate::Param* pParam = this->GetSteelValue(type , section);
		if(NULL == pParam) return res;

		const double E = pParam->E;
		const double N = pParam->N;
		const double P = pParam->P;

		res = 2*E + N*P;
	}

	return res;
}

/**
	@brief	send command to MSTN J
	@author	humkyung
	@date	2013.06.26
*/
void CSmartSteelDoc::SendKeyIn(const STRING_T& keyin, HWND fwpWnd, CString findCaption)
{
	for(int nIndex=0; nIndex < int(keyin.size()); ++nIndex )
	{
		SHORT ch = VkKeyScan(keyin[nIndex]);
		if(1 == HIBYTE(ch))
		{
			/// shift key press
			keybd_event(VK_SHIFT , MapVirtualKey(VK_SHIFT, 0) , 0, 0 );
		}
		keybd_event(ch , MapVirtualKey(keyin[nIndex] , 0) , 0 , 0 );	/// 0xE0 
		keybd_event(ch , MapVirtualKey(keyin[nIndex] , 0) , KEYEVENTF_KEYUP, 0 );
		if(1 == HIBYTE(ch))
		{
			/// shift key up
			keybd_event(VK_SHIFT , MapVirtualKey(VK_SHIFT, 0) , KEYEVENTF_KEYUP , 0 );
		}
  	}
}

/**
	@brief	send command to MSTN J with parameter given data file path

	@author	humkyung

	@date	2013.06.26
*/
bool CSmartSteelDoc::SendCommand(const STRING_T& sDatFilePath , const STRING_T& sConfigFilePath)
{
	CString tmpcmd;
	/// 명령 만들기
	HWND fwpWnd = FindWindow( _T("MstnTop") , NULL);
	///프레임웤이 열려있지 않다면 종료
	if(NULL == fwpWnd)
	{
		AfxMessageBox(_T("Microstation isn't open") , MB_OK , MB_ICONEXCLAMATION);
		return false;
	}
	HWND hKeyinWnd = FindWindowEx(fwpWnd, NULL , _T("MStnChild") , _T("Command Window") );
	if(NULL == hKeyinWnd)
	{
		AfxMessageBox(_T("Please, open Command Window and then retry.") , MB_OK , MB_ICONEXCLAMATION);
		return false;
	}

	AttachThreadInput(GetCurrentThreadId(), GetWindowThreadProcessId(fwpWnd,NULL),TRUE);
	SetForegroundWindow(fwpWnd);

	/// clear command
	keybd_event(VK_ESCAPE, MapVirtualKey(VK_ESCAPE , 0), 0, 0);
	keybd_event(VK_ESCAPE, MapVirtualKey(VK_ESCAPE , 0), KEYEVENTF_KEYUP, 0);
	Sleep(20);

	SendKeyIn(_T("mdl l FWPTool HK100 ") + sDatFilePath + _T(" ") + sConfigFilePath);
	Sleep(100);
	keybd_event(VK_RETURN, MapVirtualKey(VK_RETURN , 0), 0, 0);
	keybd_event(VK_RETURN, MapVirtualKey(VK_RETURN , 0), KEYEVENTF_KEYUP, 0);
	
	////////////////////////////
	AttachThreadInput(GetCurrentThreadId(), GetWindowThreadProcessId(fwpWnd,NULL),FALSE);

	return true;
}

#ifdef	SMART_STEEL
/******************************************************************************
    @brief		export model to BREP
	@author     humkyung
    @date       2013-07-19
    @class      CSmartSteelDoc
    @function   OnExportBREP
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnExportBREP()
{
	CImportExport::SaveBREP(this->GetAISContext());
}
#endif

/**
	@brief
	@author	humkyung
	@date	2013.08.04
*/
void CSmartSteelDoc::OnUpdateAppExport(CCmdUI* pCmdUI)
{
	pCmdUI->Enable(!m_oGussetPlateList.empty() || !m_oEndPlateList.empty());
}

/******************************************************************************
    @brief		export gusset/end plate
	@author     humkyung
    @date       2013-06-25
    @class      CSmartSteelDoc
    @function   OnAppExport
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnAppExport()
{
	if(!m_oGussetPlateList.empty() || !m_oEndPlateList.empty())
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		CString sTempPath = CFileTools::GetTempPath();
		if(!sTempPath.IsEmpty() && (_T('\\') != sTempPath.GetAt(sTempPath.GetLength() - 1)))
		{
			sTempPath += _T("\\");
		}
		OFSTREAM_T ofile(sTempPath + _T("plate.mdt"));
		if(ofile.is_open())
		{
			const double dScale = (UNIT::M == docData.m_oPlateCfg.unit_) ? (1000.0) : 1.;

			ofile.precision( 3 );									/// 소수점 자릿수 3
			ofile.setf(ios_base:: fixed, ios_base:: floatfield);	/// 소수점 방식으로 표시됨
			for(list<CGussetPlate*>::iterator itr = m_oGussetPlateList.begin();itr != m_oGussetPlateList.end();++itr)
			{
				(*itr)->Write(ofile , dScale);
			}
			for(list<CEndPlate*>::iterator itr = m_oEndPlateList.begin();itr != m_oEndPlateList.end();++itr)
			{
				(*itr)->Write(ofile , dScale);
			}
			ofile.close();

			OFSTREAM_T ocfg(sTempPath + _T("plate.cfg"));
			if(ocfg.is_open())
			{
				ocfg << _T("0,0,0,") << docData.m_oPlateCfg.class_ << _T(",") << docData.m_oPlateCfg.grade_;
				ocfg.close();
			}

			SendCommand((sTempPath + _T("plate.mdt")).operator LPCTSTR() , (sTempPath + _T("plate.cfg")).operator LPCTSTR());
		}
	}
}

/******************************************************************************
    @brief		edit plate options
	@author     humkyung
    @date       2013-07-01
    @class      CSmartSteelDoc
    @function   OnAppOptions
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnAppOptions()
{
	CAppDocData& docData = CAppDocData::GetInstance();
#ifdef	SMART_STEEL
	CSmartSteelPropertySheet dlg(0 , m_sIniFilePath);
	if(IDOK == dlg.DoModal())
	{
	}
#else
	CPlateConfigDlg dlg(m_sIniFilePath);
	dlg.DoModal();
#endif
}

/**
	@brief
	@author	humkyung
	@date	2013.08.18
*/
void CSmartSteelDoc::OnUpdateOCCViewShaded(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(AIS_Shaded == m_eDisplayMode);
}

/******************************************************************************
    @brief		view linear member and gusset/end plate as shaded
	@author     humkyung
    @date       2013-07-05
    @class      CSmartSteelDoc
    @function   OnOCCViewShaded
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnOCCViewShaded()
{
	for(list<OCC::CComplexShapeEntity* >::iterator itr = m_oSteelShapeList.begin();itr != m_oSteelShapeList.end();++itr)
	{
		(*itr)->Show(this->GetAISContext() , true);
	}

	POSITION pos = this->GetFirstViewPosition();
	while(pos)
	{
		CView* pView = this->GetNextView(pos);
		if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
		{
			CSmartSteelView* _pView = static_cast<CSmartSteelView*>(pView);
			if(NULL != _pView) _pView->RedrawWindow();
		}
	}

	m_eDisplayMode = AIS_Shaded;
}

/**
	@brief
	@author	humkyung
	@date	2013.08.18
*/
void CSmartSteelDoc::OnUpdateOCCViewWireFrame(CCmdUI* pCmdUI)
{
	pCmdUI->SetCheck(AIS_WireFrame == m_eDisplayMode);
}

/******************************************************************************
    @brief		view linear member as wireframe and gusset/end plate as shaded
	@author     humkyung
    @date       2014-08-18
    @class      CSmartSteelDoc
    @function   OnOCCViewWireframe
    @return     void
******************************************************************************/
void CSmartSteelDoc::OnOCCViewWireframe()
{
	for(list<OCC::CComplexShapeEntity* >::iterator itr = m_oSteelShapeList.begin();itr != m_oSteelShapeList.end();++itr)
	{
		(*itr)->Show(this->GetAISContext() , false);
	}

	POSITION pos = this->GetFirstViewPosition();
	while(pos)
	{
		CView* pView = this->GetNextView(pos);
		if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
		{
			CSmartSteelView* _pView = static_cast<CSmartSteelView*>(pView);
			if(NULL != _pView) _pView->RedrawWindow();
		}
	}

	m_eDisplayMode = AIS_WireFrame;
}

/******************************************************************************
    @author     humkyung
    @date       2013-07-25
    @class      CSmartSteelDoc
    @function   ZoomOCCEntity
    @return     int
    @brief		zoom given entity
******************************************************************************/
int CSmartSteelDoc::ZoomOCCEntity(CSteelPlate* pPlate)
{
	assert(pPlate && "pPlate is NULL");

	Bnd_Box oBndBox;
	if(pPlate && pPlate->GetShapeEntList())
	{
		vector<OCC::CComplexShapeEntity*>* pShapeEntList = pPlate->GetShapeEntList();
		if(NULL == pShapeEntList) return ERROR_BAD_ENVIRONMENT;
		for(vector<OCC::CComplexShapeEntity*>::iterator itr = pShapeEntList->begin();itr != pShapeEntList->end();++itr)
		{
			oBndBox.Add((*itr)->BoundingBox());
		}
		if(oBndBox.IsVoid() || oBndBox.IsWhole() || oBndBox.IsThin(std::numeric_limits<double>::epsilon())) ERROR_BAD_ENVIRONMENT;

		Standard_Real aXmin = 0, aYmin = 0 , aZmin = 0 , aXmax = 0 , aYmax = 0 , aZmax = 0;
		oBndBox.Get(aXmin , aYmin , aZmin , aXmax , aYmax , aZmax);
		Standard_Real size = max(max(aXmax - aXmin , aYmax - aYmin) , aZmax - aZmin);
		oBndBox.Enlarge(size);

		pPlate->DrawDimension(GetAISContext());

		POSITION pos = this->GetFirstViewPosition();
		while(pos)
		{
			CView* pView = this->GetNextView(pos);
			if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
			{
				CSmartSteelView* _pView = static_cast<CSmartSteelView*>(pView);

				_pView->ZoomWin( oBndBox );
			}
		}
	}
	else
	{
		CMainFrame* pFrameWnd = CMainFrame::GetInstance();
		pFrameWnd->SendMessage(DISPLAY_STATUSBAR , WPARAM(_T("Fail to zoom entity")) , 0);
	}
		
	return ERROR_SUCCESS;
}

/******************************************************************************
	@brief		zoom given entity
	@author     humkyung
    @date       2014-12-16
    @class      CSmartSteelDoc
    @function   ZoomOCCEntity
    @return     int
******************************************************************************/
int CSmartSteelDoc::ZoomOCCEntity(CSDNFLinearMember* pMember)
{
	assert(pMember && "pMember is NULL");

	Bnd_Box oBndBox;
	if(pMember)
	{
		tr1::unordered_map<STRING_T,OCC::CComplexShapeEntity* >::iterator where = m_oLinearShapeMap.find(pMember->MemberID());
		if(where != m_oLinearShapeMap.end())
		{
			oBndBox.Add(where->second->BoundingBox());
			if(oBndBox.IsVoid() || oBndBox.IsWhole() || oBndBox.IsThin(std::numeric_limits<double>::epsilon())) ERROR_BAD_ENVIRONMENT;

			Standard_Real aXmin = 0, aYmin = 0 , aZmin = 0 , aXmax = 0 , aYmax = 0 , aZmax = 0;
			oBndBox.Get(aXmin , aYmin , aZmin , aXmax , aYmax , aZmax);
			Standard_Real size = max(max(aXmax - aXmin , aYmax - aYmin) , aZmax - aZmin);
			oBndBox.Enlarge(size);

			POSITION pos = this->GetFirstViewPosition();
			while(pos)
			{
				CView* pView = this->GetNextView(pos);
				if(pView->IsKindOf(RUNTIME_CLASS(CSmartSteelView)))
				{
					CSmartSteelView* _pView = static_cast<CSmartSteelView*>(pView);
					_pView->ZoomWin( oBndBox );
				}
			}
		}
	}
	else
	{
		CMainFrame* pFrameWnd = CMainFrame::GetInstance();
		pFrameWnd->SendMessage(DISPLAY_STATUSBAR , WPARAM(_T("Fail to zoom entity")) , 0);
	}
		
	return ERROR_SUCCESS;
}


/**
	@brief	return shape parameter corresponding to given section name
	@author	humkyung
	@date	2013.06.28
*/
CSteelSectionBuilder::ShapeParam* CSmartSteelDoc::GetShapeParamOf(const STRING_T& sSectionName) const
{
	CSteelSectionBuilder::ShapeParam* res = NULL;

	STRING_T tmp(sSectionName);
	if('"' == tmp[0]) tmp = tmp.substr(1);
	if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);
	tmp = CString(tmp.c_str()).MakeUpper().operator LPCTSTR();

	map<STRING_T , CSteelSectionBuilder::ShapeParam* >::const_iterator where = m_oShapeParamMap.find(tmp);
	if(where != m_oShapeParamMap.end())
	{
		res = where->second;
	}

	return res;
}

/**
	@brief	return end plate parameter corresponding to given section name

	@author	humkyung

	@date	2013.07.04
*/
CEndPlate::Param* CSmartSteelDoc::GetEndPlateParamOf(const STRING_T& sSectionName) const
{
	CEndPlate::Param* res = NULL;

	STRING_T tmp(sSectionName);
	if('"' == tmp[0]) tmp = tmp.substr(1);
	if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);
	map<STRING_T , CEndPlate::Param* >::const_iterator where = m_oEndPlateParamMap.find(tmp);
	if(where != m_oEndPlateParamMap.end())
	{
		res = where->second;
	}

	return res;
}

/**
	@brief	return plate thickness
	@author	humkyung
	@date	2013.06.26
*/
double CSmartSteelDoc::GetPlateThicknessOf(const STRING_T& sSectionName) const
{
	double res = 15.0;	/// default value is 15

	CAppDocData& docData = CAppDocData::GetInstance();
	res *= (UNIT::M == docData.m_oPlateCfg.unit_) ? (1.0/1000.0) : 1.;

	STRING_T tmp(sSectionName);
	if('"' == tmp[0]) tmp = tmp.substr(1);
	if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);
	map<STRING_T , CGussetPlate::Param* >::const_iterator where = m_oGussetPlateParamMap.find(tmp);
	if(where != m_oGussetPlateParamMap.end())
	{
		res = where->second->T;
		return res;
	}

	return res;
}

/**
	@brief	return color of linear member

	@author	humkyung

	@date	2013.06.27
*/
STRING_T CSmartSteelDoc::GetColorStringOf(const CSDNFLinearMember::ElmType& type) const
{
	STRING_T res(_T("0,0,0"));
	switch(type)
	{
		case CSDNFLinearMember::COLUMN:
			res = _T("255,0,0");
			break;
		case CSDNFLinearMember::BEAM:
			res = _T("255,255,0");
			break;
		case CSDNFLinearMember::HBRACE:
			res = _T("0,0,255");
			break;
		case CSDNFLinearMember::VBRACE:
			res = _T("0,255,0");
			break;
		default:
			break;
	}

	return res;
}