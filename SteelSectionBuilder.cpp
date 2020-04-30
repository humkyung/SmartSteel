#include "StdAfx.h"
#include <assert.h>
#include <SDNFColumn.h>
#include "SmartSteelDoc.h"
#include "SteelSectionBuilder.h"

CSteelSectionBuilder::CSteelSectionBuilder(void)
{
}

CSteelSectionBuilder::~CSteelSectionBuilder(void)
{
}

/**
	@brief	build shape for linear member
	@author	humkyung
	@date	2013.06.21
*/
int CSteelSectionBuilder::Build(CSDNFLinearMember* pMember , CSmartSteelDoc* pDoc)
{
	assert(pMember && pDoc && "pMember or pDoc is NULL");

	if(pMember && pDoc)
	{
		m_oSectionPntList.clear();

		STRING_T section = pMember->section();
		STRING_T tmp(section);
		if(tmp.empty()) return ERROR_BAD_ENVIRONMENT;	/// check section name - 2013.10.25 added by humkyung
		if('"' == tmp[0]) tmp = tmp.substr(1);
		if('"' == tmp[tmp.length() - 1]) tmp = tmp.substr(0 , tmp.length() - 1);

		CSteelSectionBuilder::ShapeParam* pShapeParam = NULL;
		map<STRING_T , CSteelSectionBuilder::ShapeParam* >::iterator where = pDoc->m_oShapeParamMap.find(tmp);
		if(where == pDoc->m_oShapeParamMap.end()) return ERROR_BAD_ENVIRONMENT;
		pShapeParam = where->second;
		
		CIsVect3d OV = pMember->OriVector();
		CIsVect3d U  = CIsVect3d(pMember->end() - pMember->start()).Normalize();
		CIsVect3d W(U*OV) , V(W*U);
		W.Normalize();
		V.Normalize();
		if((NULL != pShapeParam ) && (_T("WFB") == pShapeParam->Shape))
		{			
			m_oSectionPntList.push_back(CIsPoint3d(-V*pShapeParam->Height*0.5 + -W*pShapeParam->Width*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[0] + W*pShapeParam->Width);
			m_oSectionPntList.push_back(m_oSectionPntList[1] + V*pShapeParam->t2);
			m_oSectionPntList.push_back(m_oSectionPntList[2] - W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[3] + V*(pShapeParam->Height - pShapeParam->t2*2));
			m_oSectionPntList.push_back(m_oSectionPntList[4] + W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[5] + V*(pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[6] - W*(pShapeParam->Width));
			m_oSectionPntList.push_back(m_oSectionPntList[7] - V*(pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[8] + W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[9] - V*(pShapeParam->Height - pShapeParam->t2*2));
			m_oSectionPntList.push_back(m_oSectionPntList[10] - W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[11] - V*(pShapeParam->t2));
		}
		else if((NULL != pShapeParam ) && (_T("CHANNEL") == pShapeParam->Shape))
		{
			m_oSectionPntList.push_back(CIsPoint3d(-V*pShapeParam->Height*0.5 - W*pShapeParam->Width*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[0] + W*pShapeParam->Width);
			m_oSectionPntList.push_back(m_oSectionPntList[1] + V*pShapeParam->t2);
			m_oSectionPntList.push_back(m_oSectionPntList[2] - W*(pShapeParam->Width - pShapeParam->t1));
			m_oSectionPntList.push_back(m_oSectionPntList[3] + V*(pShapeParam->Height - pShapeParam->t2*2));
			m_oSectionPntList.push_back(m_oSectionPntList[4] + W*(pShapeParam->Width - pShapeParam->t1));
			m_oSectionPntList.push_back(m_oSectionPntList[5] + V*pShapeParam->t2);
			m_oSectionPntList.push_back(m_oSectionPntList[6] - W*pShapeParam->Width);
			m_oSectionPntList.push_back(m_oSectionPntList[7] - V*pShapeParam->Height);
		}
		else if((NULL != pShapeParam ) && (_T("TEE") == pShapeParam->Shape))
		{
			m_oSectionPntList.push_back(CIsPoint3d(-V*pShapeParam->Height*0.5 - W*pShapeParam->t1*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[0] + W*pShapeParam->t1);
			m_oSectionPntList.push_back(m_oSectionPntList[1] + V*(pShapeParam->Height - pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[2] + W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[3] + V*(pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[4] - W*(pShapeParam->Width));
			m_oSectionPntList.push_back(m_oSectionPntList[5] - V*(pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[6] + W*((pShapeParam->Width - pShapeParam->t1)*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[7] - V*(pShapeParam->Height - pShapeParam->t2));
		}
		else if((NULL != pShapeParam ) && (_T("ANGLE") == pShapeParam->Shape))
		{
			m_oSectionPntList.push_back(CIsPoint3d(-V*pShapeParam->Height*0.5 - W*pShapeParam->Width*0.5));
			m_oSectionPntList.push_back(m_oSectionPntList[0] + W*pShapeParam->Width);
			m_oSectionPntList.push_back(m_oSectionPntList[1] + V*(pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[2] - W*(pShapeParam->Width - pShapeParam->t1));
			m_oSectionPntList.push_back(m_oSectionPntList[3] + V*(pShapeParam->Height - pShapeParam->t2));
			m_oSectionPntList.push_back(m_oSectionPntList[4] - W*(pShapeParam->t1));
			//m_oSectionPntList.push_back(m_oSectionPntList[5] - V*(pShapeParam->Height));
		}
		AdjustCardinalPoint(pMember , W , pShapeParam->Width , V , pShapeParam->Height);
		RotateSection(U , pMember->rotation());

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	rotate section shape about rotation
	@author	humkyung
	@date	2013.06.24
*/
int CSteelSectionBuilder::RotateSection(const CIsVect3d& norm , const double& dRotation)
{
	const double dRad = DEG2RAD(dRotation);
	CIsQuat quat(norm , dRad);
	for(vector<CIsPoint3d>::iterator itr = m_oSectionPntList.begin();itr != m_oSectionPntList.end();++itr)
	{
		(*itr) = itr->Rotate(quat);
	}

	return ERROR_SUCCESS;
}

/**
	@brief	adjust cardinal point
	@author	humkyung
	@date	2013.06.24
*/
int CSteelSectionBuilder::AdjustCardinalPoint(CSDNFLinearMember* pMember , const CIsVect3d& W , const double& dWidth , const CIsVect3d& V , const double& dHeight)
{
	assert(pMember && "pMember is NULL");
	if(pMember)
	{
		CIsVect3d move(0.0 , 0.0 , 0.0);
		switch(pMember->CardinalPoint())
		{
			case 1:
				move = -W*dWidth*0.5 - V*dHeight*0.5;
				break;
			case 2:
				move = -V*dHeight*0.5;
				break;
			case 3:
				move = W*dWidth*0.5 - V*dHeight*0.5;
				break;
			case 4:
				move = -W*dWidth*0.5;
				break;
			case 6:
				move = W*dWidth*0.5;
				break;
			case 7:
				move = -W*dWidth*0.5 + V*dHeight*0.5;
				break;
			case 8:
				move = V*dHeight*0.5;
				break;
			case 9:
				move = W*dWidth*0.5 + V*dHeight*0.5;
				break;
		}
		for(vector<CIsPoint3d>::iterator itr = m_oSectionPntList.begin();itr != m_oSectionPntList.end();++itr)
		{
			(*itr) -= move;
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}