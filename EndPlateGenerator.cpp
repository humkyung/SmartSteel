#include "StdAfx.h"
#include <assert.h>
#include "SmartSteelDoc.h"
#include "EndPlateGenerator.h"
#include "AppDocData.h"
#include "MainFrm.h"

CEndPlateGenerator::CEndPlateGenerator(void)
{
}

CEndPlateGenerator::~CEndPlateGenerator(void)
{
}

/**
	@brief	return the instance of CEndPlateGenerator

	@author	humkyung

	@date	2013.06.24
*/
CEndPlateGenerator& CEndPlateGenerator::GetInstance(void)
{
	static CEndPlateGenerator __instance__;

	return __instance__;
}

/**
	@brief	check given dir is parallel to width dir of linear member

	@auhor	humkyung

	@date	2013.07.04
*/
bool CEndPlateGenerator::IsWidthDir(CSDNFLinearMember* pMember , const CIsVect3d& dir) const
{
	assert(pMember && "pMember is NULL");

	if(pMember)
	{
		CIsVect3d W = pMember->W();
		CIsQuat quat(pMember->Direction() , DEG2RAD(pMember->rotation()));
		W = W.Rotate(quat);
		return !equals((W.DotProduct(dir)) , 0.0);
	}

	return false;
}

/**
	@brief	check given dir is parallel to height dir of linear member

	@auhor	humkyung

	@date	2013.07.04
*/
bool CEndPlateGenerator::IsHeightDir(CSDNFLinearMember* pMember , const CIsVect3d& dir) const
{
	assert(pMember && "pMember is NULL");

	if(pMember)
	{
		CIsVect3d V = pMember->V();
		CIsQuat quat(pMember->Direction() , DEG2RAD(pMember->rotation()));
		V = V.Rotate(quat);
		return !equals((V.DotProduct(dir)) , 0.0);
	}

	return false;
}

/**
	@brief	generate end plate
	@auhor	humkyung
	@date	2013.06.24
*/
int CEndPlateGenerator::Generate(list<CEndPlate*>& oEndPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");

	if(pConnPnt && pDoc && (CSteelConnPoint::COLUMN_TO_BEAM == pConnPnt->Type()))
	{
		if(2 == pConnPnt->GetMemberSize())
		{
			CSDNFLinearMember* pColumn = pConnPnt->GetMemberAt(0);
			CSDNFLinearMember* pBeam = pConnPnt->GetMemberAt(1);
			
			/// return if beam length is less than 2m when check beam length option is on(refer to : http://humkyung100.cafe24.com:8080/projects/SmartSteel/ticket/26) - 2013.10.24 added by humkyung
			CAppDocData& docData = CAppDocData::GetInstance();
			if(true == docData.m_oPlateCfg.generate_endplate_depend_on_beam_length)
			{
				const double dLength = pBeam->end(CSDNFElement::METER).DistanceTo(pBeam->start(CSDNFElement::METER));
				if(dLength < 2.0) return ERROR_SUCCESS;
			}
			/// up to here

			CIsVect3d column_dir   = pColumn->end(CSDNFElement::METER) - pColumn->start(CSDNFElement::METER);
			column_dir.Normalize();
			
			vector<CIsVect3d> oBeamDirList;
			/// check column meets end points of beam(ticket #7) - 2013.07.10 added by humkyung
			if(!(pConnPnt->origin() == pBeam->start(CSDNFElement::METER)) && !(pConnPnt->origin() == pBeam->end(CSDNFElement::METER)))
			{
				CIsVect3d beam_dir = -(pBeam->Direction().Normalize());
				oBeamDirList.push_back(beam_dir);
				oBeamDirList.push_back(-beam_dir);
			}
			else
			{
				CIsVect3d beam_dir = pBeam->Direction();
				if(pConnPnt->origin() == pBeam->end(CSDNFElement::METER)) beam_dir = -beam_dir;
				oBeamDirList.push_back(beam_dir.Normalize());
			}
			
			CSteelSectionBuilder::ShapeParam* pColumnShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
			CSteelSectionBuilder::ShapeParam* pBeamShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(1)->section());
			CEndPlate::Param* pEndPlateParam = pDoc->GetEndPlateParamOf(pConnPnt->GetMemberAt(1)->section());
			/// check end plate parameter(refer to http://humkyung100.cafe24.com:8080/projects/SmartSteel/ticket/25) - 2013.10.24 added by humkyung
			if((NULL == pEndPlateParam) || (!pEndPlateParam->IsValid()))
			{
				OSTRINGSTREAM_T oss;
				oss << _T("Param of EndPlate for ") << pConnPnt->GetMemberAt(1)->section() << _T(" is invalid");
				CMainFrame* pFrameWnd = CMainFrame::GetInstance();
				pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(oss.str().c_str()) , MessageType::MESSAGE_WARNING);
				
				return ERROR_BAD_ENVIRONMENT;
			}
			/// up to here
			if((NULL == pEndPlateParam) || (NULL == pColumnShapeParam) || (NULL == pBeamShapeParam)) return ERROR_BAD_ENVIRONMENT;
			
			CIsVect3d sec_dir = pBeam->V();
			for(vector<CIsVect3d>::iterator itr = oBeamDirList.begin();itr != oBeamDirList.end();++itr)
			{
				CIsPoint3d ptOrigin(pConnPnt->origin());
				ptOrigin -= sec_dir*pBeamShapeParam->Height;
				if((ptOrigin.z() < pColumn->start(CSDNFElement::METER).z()) && (ptOrigin.z() < pColumn->end(CSDNFElement::METER).z())) continue;
				if(this->IsHeightDir(pConnPnt->GetMemberAt(0) , *itr))
				{
					ptOrigin += (*itr)*pColumnShapeParam->Height*0.5;
				}
				else
				{
					/// end plate is only created on column's flange direction.
					continue;
				}
			
				CEndPlate* pEndPlate = new CEndPlate(pConnPnt);
				{
					pEndPlate->m_oSectionShapePntList.push_back(ptOrigin);
					pEndPlate->m_oSectionShapePntList.push_back(ptOrigin + (*itr)*(pEndPlateParam->M));
					pEndPlate->m_oSectionShapePntList.push_back(ptOrigin - sec_dir*(pEndPlateParam->K) + (*itr)*((UNIT::M == docData.m_oPlateCfg.unit_) ? 0.15 : 150));
					pEndPlate->m_oSectionShapePntList.push_back(ptOrigin - sec_dir*(pEndPlateParam->K));

					pEndPlate->m_norm = column_dir*(*itr);
					pEndPlate->m_norm.Normalize();
					pEndPlate->m_dThickness = pEndPlateParam->T;

					oEndPlateList.push_back(pEndPlate);
				}
			}
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}