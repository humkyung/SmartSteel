#include "StdAfx.h"
#include <assert.h>
#include <IsLine3d.h>
#include <IsPlane3d.h>
#include "GussetPlate.h"
#include "GussetPlateGenerator.h"
#include "MainFrm.h"
#include "AppDocData.h"
#include "HorGussetPlateGenerator.h"

CGussetPlateGenerator::CGussetPlateGenerator(void)
{
}

CGussetPlateGenerator::~CGussetPlateGenerator(void)
{
}

/**
	@brief	return the instance of CGussetPlateGenerator
	@author	humkyung
	@date	2013.06.24
*/
CGussetPlateGenerator& CGussetPlateGenerator::GetInstance(void)
{
	static CGussetPlateGenerator __instance__;

	return __instance__;
}

/**
	@brief	check given dir is parallel to width dir of linear member
	@auhor	humkyung
	@date	2013.06.28
*/
bool CGussetPlateGenerator::IsWebDir(CSDNFLinearMember* pMember , const CIsVect3d& dir)
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
	@date	2013.06.28
*/
bool CGussetPlateGenerator::IsFlangeDir(CSDNFLinearMember* pMember , const CIsVect3d& dir)
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
	@brief	return the offset from linear member depend on connection dir.
	@auhor	humkyung
	@date	2013.06.28
*/
double CGussetPlateGenerator::GetOffsetFromLinearMember(CSDNFLinearMember* pMember , const CSDNFLinearMember::ElmType& type , const CIsPoint3d& ptConn , const CIsVect3d& oConnDir , CSmartSteelDoc* pDoc)
{
	assert(pMember && pDoc && "pMember or pDoc is NULL");
	double res = 0.0;

	if(pMember && pDoc)
	{
		CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pMember->section());
		if(NULL != pShapeParam)
		{
			if(CGussetPlateGenerator::IsWebDir(pMember , oConnDir))		/// connected to web face
			{
				CIsLine3d line3d(pMember->start(CSDNFLinearMember::METER) , pMember->end(CSDNFLinearMember::METER));
				const double dist = line3d.DistanceTo(ptConn);

				const int iCardinalPnt = pMember->CardinalPoint();
				switch(iCardinalPnt)
				{
				case 1: case 4: case 7:
					res = (CGussetPlateGenerator::IsSameDir(pMember->W() , oConnDir)) ? (pShapeParam->Width + pShapeParam->t1) : 0.0;
					break;
				case 5:
					if(CSDNFLinearMember::HBRACE == type)
					{
						res = (pShapeParam->t1 - dist*2.0);
					}
					else
					{
						res = (CGussetPlateGenerator::IsSameDir(pMember->W() , oConnDir)) ? (pShapeParam->Width - dist) : (pShapeParam->Width + dist);
					}
					break;
				case 2: case 8:
					res = (CSDNFLinearMember::HBRACE == type) ? pShapeParam->t1 : pShapeParam->Width;
					break;
				case 3: case 6: case 9:
					res = (CGussetPlateGenerator::IsSameDir(pMember->W() , oConnDir)) ? 0.0 : (pShapeParam->Width + pShapeParam->t1);
					break;
				}
			}
			else if(CGussetPlateGenerator::IsFlangeDir(pMember , oConnDir))	/// connected to flange face
			{
				CIsLine3d line3d(pMember->start(CSDNFLinearMember::METER) , pMember->end(CSDNFLinearMember::METER));
				const double dist = line3d.DistanceTo(ptConn);

				const int iCardinalPnt = pMember->CardinalPoint();
				switch(iCardinalPnt)
				{
				case 1: case 2: case 3:
					res = (CGussetPlateGenerator::IsSameDir(pMember->V() , oConnDir)) ? (pShapeParam->Height - dist)*2.0 : 0.0;
					break;
				case 4: case 5: case 6:
					res = (CGussetPlateGenerator::IsSameDir(pMember->V() , oConnDir)) ? (pShapeParam->Height - dist) : (pShapeParam->Height + dist);
					break;
				case 7: case 8: case 9:
					res = (CGussetPlateGenerator::IsSameDir(pMember->V() , oConnDir)) ? dist*2.0 : (pShapeParam->Height - dist)*2.0;
					break;
				}
			}
		}
	}

	return res;
}

/**
	@brief	check given two direction has same direction

	@auhor	humkyung

	@date	2013.06.28
*/
bool CGussetPlateGenerator::IsSameDir(const CIsVect3d lhs , const CIsVect3d& rhs)
{
	return equals(lhs.DotProduct(rhs) , 1.0);
}

/**
	@brief	generate gusset plate for COLUMN TO VER. BRACE
	@auhor	humkyung
	@date	2013.06.28
*/
int CGussetPlateGenerator::Generate4ColumnToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");

	if(pConnPnt && pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		try
		{
			if(2 == pConnPnt->GetMemberSize())
			{
				CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);

				CIsVect3d column_dir = pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER);
				column_dir.Normalize();
				CIsVect3d vbrace_dir = pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER);
				if(pConnPnt->origin() == pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER)) vbrace_dir = -vbrace_dir;
				vbrace_dir.Normalize();
				pGussetPlate->m_norm = pConnPnt->GetMemberAt(1)->W();
				pGussetPlate->m_norm.Normalize();
				CIsVect3d beam_dir = pGussetPlate->m_norm*column_dir;
				beam_dir.Normalize();
				
				/// 2013.09.02 added by humkyung
				double dot = vbrace_dir.DotProduct(beam_dir);
				if(equals(dot , 0.0))
				{
					pGussetPlate->m_norm = pConnPnt->GetMemberAt(1)->V();
					pGussetPlate->m_norm.Normalize();
					beam_dir = pGussetPlate->m_norm*column_dir;
					beam_dir.Normalize();
				}
				/// up to here

				dot = vbrace_dir.DotProduct(beam_dir);
				if(dot < 0.0)
				{
					beam_dir = -beam_dir;
				}

				const double alpha = acos(beam_dir.DotProduct(vbrace_dir));

				CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
				const double B = this->GetOffsetFromLinearMember(pConnPnt->GetMemberAt(0) , pConnPnt->GetMemberAt(1)->Type() , pConnPnt->origin() , beam_dir , pDoc);///(NULL != pShapeParam) ? pShapeParam->t1 : 0.0;
				CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
				double A = (NULL != pParam) ? pParam->A : -1;
				double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
				if((-1 == A) || (-1 == jointL))
				{
					SAFE_DELETE(pGussetPlate);
					return ERROR_BAD_ENVIRONMENT;
				}
				double totalL = B*0.5*(1/cos(alpha)) + (SHAPE_LEN_OFFSET)*(1/cos(alpha)) + A*sin(alpha) + jointL;
				CIsVect3d cross = vbrace_dir*pGussetPlate->m_norm;
				cross.Normalize();

				{
					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + beam_dir*(B*0.5));
					pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + (beam_dir*(SHAPE_START_LEN)));

					CIsPoint3d tmp[2];
					tmp[0] = pConnPnt->origin() + vbrace_dir*totalL + cross*A*0.5;
					tmp[1] = pConnPnt->origin() + vbrace_dir*totalL - cross*A*0.5;
					double d[2]={0,0};
					d[0] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[0]);
					d[1] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[1]);
					if(d[0] < d[1])
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
					}
					else
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
					}

					{
						CIsVect3d dir(pGussetPlate->m_oSectionShapePntList[pGussetPlate->m_oSectionShapePntList.size()-1] - pGussetPlate->m_oSectionShapePntList[0]);
						dot = dir.DotProduct(column_dir);
						pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + column_dir*dot);
					}
					pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(1)->section());

					oGussetPlateList.push_back(pGussetPlate);
				}

				return ERROR_SUCCESS;
			}
		}
		catch(...)
		{
		}
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	generate gusset plate for BEAM TO HOR. BRACE
	@auhor	humkyung
	@date	2013.06.28
*/
int CGussetPlateGenerator::Generate4BeamToHBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*>& oMemList , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc or pMem is NULL");

	if(pConnPnt && pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();
		try
		{
			if(2 == oMemList.size())
			{
				CIsVect3d beam_dir   = oMemList[0]->Direction();
				beam_dir.Normalize();
				CIsVect3d hbrace_dir = oMemList[1]->Direction();
				if(pConnPnt->origin().DistanceTo(oMemList[1]->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(oMemList[1]->start(CSDNFElement::METER)))
				{
					hbrace_dir = -hbrace_dir;
				}
				hbrace_dir.Normalize();
				const double angle = (beam_dir.AngleTo(hbrace_dir));
				beam_dir = (beam_dir.DotProduct(hbrace_dir) < 0) ? -beam_dir : beam_dir;	/// check the angle between beam and hbrace - 2013.06.26 added by humkyung
				if(!equals(fabs(angle) , PI*0.5))	/// if not angle between two vectors isn't 90 degree
				{
					CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
					pGussetPlate->m_norm = beam_dir*hbrace_dir;
					pGussetPlate->m_norm.Normalize();
					/// force to locate gusset plate over linear member - 2013.07.28 added by humkyung
					if(pGussetPlate->m_norm.dz() < 0.0) pGussetPlate->m_norm.Set(-pGussetPlate->m_norm.dx() , -pGussetPlate->m_norm.dy() , -pGussetPlate->m_norm.dz());
					CIsVect3d base_dir = pGussetPlate->m_norm*beam_dir;
					base_dir.Normalize();
					if(base_dir.DotProduct(hbrace_dir) < 0.0) base_dir.Set(-base_dir.dx() , -base_dir.dy() , -base_dir.dz());

					const double alpha = acos(base_dir.DotProduct(hbrace_dir));

					CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(oMemList[0]->section());
					const double offset = this->GetOffsetFromLinearMember(oMemList[0] , oMemList[1]->Type() , pConnPnt->origin() , base_dir , pDoc);
					CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , oMemList[1]->section());
					const double A = (NULL != pParam) ? pParam->A : -1;
					double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , oMemList[1]->section());
					if((-1 == A) || (-1 == jointL))
					{
						SAFE_DELETE(pGussetPlate);
						return ERROR_BAD_ENVIRONMENT;
					}
					
					const double B = (IsFlangeDir(pConnPnt->GetMemberAt(0) , hbrace_dir)) ? pShapeParam->Height : pShapeParam->Width;
					const double totalL = B*0.5*(1/cos(alpha)) + (SHAPE_LEN_OFFSET)*(1/cos(alpha)) + A*sin(alpha) + jointL;
					CIsVect3d cross = hbrace_dir*pGussetPlate->m_norm;
					cross.Normalize();

					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + base_dir*(offset*0.5) - beam_dir*(SHAPE_START_LEN));
					pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + (base_dir*(B*0.5 + SHAPE_START_LEN)));

					CIsPoint3d tmp[2];
					tmp[0] = pConnPnt->origin() + hbrace_dir*totalL + cross*A*0.5;
					tmp[1] = pConnPnt->origin() + hbrace_dir*totalL - cross*A*0.5;
					double d[2]={0,0};
					d[0] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[0]);
					d[1] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[1]);
					if(d[0] < d[1])
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
					}
					else
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
					}

					CIsVect3d dirV(pGussetPlate->m_oSectionShapePntList[pGussetPlate->m_oSectionShapePntList.size()-1] - pGussetPlate->m_oSectionShapePntList[0]);
					const double len = beam_dir.DotProduct(dirV);
					pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + beam_dir*len);

					pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(oMemList[1]->section());

					oGussetPlateList.push_back(pGussetPlate);
				}
			}
			else
			{
				CIsVect3d beam_dir = oMemList[0]->Direction();
				
				CHorGussetPlateGenerator generator;
				CIsVect3d norm;
				vector<CIsPoint3d> oShapePntList;
				generator.GetOuterPntListOfPlate(oShapePntList , norm , pConnPnt , oMemList , pDoc);
				if(oShapePntList.size() >= 2)
				{
					CIsVect3d dir(oShapePntList[0] - pConnPnt->origin());
					double len = beam_dir.DotProduct(dir);
					oShapePntList.insert(oShapePntList.begin() , pConnPnt->origin() + beam_dir*len);

					dir = oShapePntList[oShapePntList.size() - 1] - pConnPnt->origin();
					len = beam_dir.DotProduct(dir);
					oShapePntList.push_back(pConnPnt->origin() + beam_dir*len);

					CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
					{
						pGussetPlate->m_oSectionShapePntList.insert(pGussetPlate->m_oSectionShapePntList.begin() , oShapePntList.begin() , oShapePntList.end());
						pGussetPlate->m_norm = norm;
						pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(1)->section());
					}
					oGussetPlateList.push_back(pGussetPlate);
				}
			}
		}
		catch(...)
		{
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	generate gusset plate for BEAM TO VER. BRACE
	@auhor	humkyung
	@date	2013.06.28
*/
int CGussetPlateGenerator::Generate4BeamToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");

	if(pConnPnt && pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		try
		{
			if(2 == pConnPnt->GetMemberSize())
			{
				CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);

				CSDNFLinearMember* pBeam = pConnPnt->GetMemberAt(0);
				CIsVect3d beam_dir   = pBeam->end(CSDNFElement::METER) - pBeam->start(CSDNFElement::METER);
				beam_dir.Normalize();
				CSDNFLinearMember* pVBrace = pConnPnt->GetMemberAt(1);
				CIsVect3d vbrace_dir = pVBrace->Direction();
				if(pConnPnt->origin().DistanceTo(pVBrace->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pVBrace->start(CSDNFElement::METER)))
				{
					vbrace_dir = -vbrace_dir;
				}
				vbrace_dir.Normalize();
				const double angle = (beam_dir.AngleTo(vbrace_dir));
				beam_dir = ((angle >= 0.0) && (angle < PI*0.5)) ? beam_dir : -beam_dir;	/// check the angle between beam and hbrace - 2013.06.26 added by humkyung
				if(!equals(fabs(angle) , PI*0.5))	/// if not angle between two vectors isn't 90 degree
				{
					pGussetPlate->m_norm = (beam_dir*vbrace_dir).Normalize();///pVBrace->W();
					pGussetPlate->m_norm.Normalize();
					CIsVect3d base_dir = pGussetPlate->m_norm*beam_dir;
					base_dir.Normalize();
					/// 2013.09.02 added by humkyung
					double dot = base_dir.DotProduct(vbrace_dir);
					if(equals(dot , 0.0))
					{
						pGussetPlate->m_norm = (beam_dir*vbrace_dir).Normalize();///pVBrace->V();
						pGussetPlate->m_norm.Normalize();
						base_dir = pGussetPlate->m_norm*beam_dir;
						base_dir.Normalize();
					}
					/// up to here

					if(base_dir.DotProduct(vbrace_dir) < 0.0)
					{
						base_dir = -base_dir;
					}

					const double alpha = acos(base_dir.DotProduct(vbrace_dir));

					CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
					const double B = this->GetOffsetFromLinearMember(pConnPnt->GetMemberAt(0) , pConnPnt->GetMemberAt(1)->Type() , pConnPnt->origin() , base_dir , pDoc);
					CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
					double A = (NULL != pParam) ? pParam->A : -1;
					double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
					if((-1 == A) || (-1 == jointL))
					{
						SAFE_DELETE(pGussetPlate);
						return ERROR_BAD_ENVIRONMENT;
					}

					double totalL = B*0.5*(1/cos(alpha)) + (SHAPE_LEN_OFFSET)*(1/cos(alpha)) + A*sin(alpha) + jointL;
					CIsVect3d cross = vbrace_dir*pGussetPlate->m_norm;
					cross.Normalize();

					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + base_dir*(B*0.5));
					pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + (base_dir*(SHAPE_START_LEN)));

					CIsPoint3d tmp[2];
					tmp[0] = pConnPnt->origin() + vbrace_dir*totalL + cross*A*0.5;
					tmp[1] = pConnPnt->origin() + vbrace_dir*totalL - cross*A*0.5;
					double d[2]={0,0};
					d[0] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[0]);
					d[1] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[1]);
					if(d[0] < d[1])
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
					}
					else
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
					}

					CIsVect3d dir = pGussetPlate->m_oSectionShapePntList[pGussetPlate->m_oSectionShapePntList.size()-1] - pGussetPlate->m_oSectionShapePntList[0];
					dot = beam_dir.DotProduct(dir);
					pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + dot*beam_dir);
					pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(1)->section());

					oGussetPlateList.push_back(pGussetPlate);
				}
			}
			else if(3 == pConnPnt->GetMemberSize())	/// BEAM , VER. BRACE , VER. BRACE
			{
				CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);

				CIsVect3d beam_dir   = pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER);
				beam_dir.Normalize();
				
				CIsVect3d vbrace_dir = pConnPnt->GetMemberAt(1)->Direction();
				if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER)))
				{
					vbrace_dir = -vbrace_dir;
				}
				vbrace_dir.Normalize();
				double angle = (beam_dir.AngleTo(vbrace_dir));
				beam_dir = ((angle >= 0.0) && (angle < PI*0.5)) ? beam_dir : -beam_dir;	/// check the angle between beam and hbrace - 2013.06.26 added by humkyung
				CIsPoint3d tmp[4];
				//if(!equals(fabs(angle) , PI*0.5))	/// if angle between two vectors isn't 90 degree
				{
					/*vector<CIsPoint3d> pts;
					{
						pts.push_back(pConnPnt->GetMemberAt(0)->start());
						pts.push_back(pConnPnt->GetMemberAt(0)->end());
						pts.push_back((pConnPnt->GetMemberAt(1)->start() + pConnPnt->GetMemberAt(1)->end())*0.5);
					}
					CIsPlane3d plane;
					CIsPlane3d::CreateOf(plane , pts);*/
					pGussetPlate->m_norm = (pConnPnt->GetMemberAt(1)->Direction()*pConnPnt->GetMemberAt(2)->Direction()).Normalize();
					/// up to here
					//vbrace_dir = vbrace_dir - pGussetPlate->m_norm.DotProduct(vbrace_dir)*pGussetPlate->m_norm;	/// force vbrace_dir to place on plane - 2014.02.11 added by humkyung

					const double alpha = acos(beam_dir.DotProduct(vbrace_dir));

					CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
					if(NULL == pShapeParam)
					{
						SAFE_DELETE(pGussetPlate);
						return ERROR_BAD_ENVIRONMENT;
					}
					double B = 0.0;
					double dot = vbrace_dir.DotProduct(beam_dir);
					CIsVect3d V = vbrace_dir - beam_dir.Normalize()*dot;
					CIsPoint3d ptTmpOrigin(pConnPnt->origin());
					if(this->IsFlangeDir(pConnPnt->GetMemberAt(0) , vbrace_dir))
					{
						CIsVect3d tmpVec = ptTmpOrigin - pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER);
						double dot = fabs(tmpVec.DotProduct(beam_dir));
						CIsPoint3d T(pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER) + beam_dir*dot);
						const double L = fabs(CIsVect3d(T - ptTmpOrigin).DotProduct(V.Normalize()));

						B = (pShapeParam->Height - L)*2;
						ptTmpOrigin += V.Normalize()*(pShapeParam->Height - L);
					}
					else if(this->IsWebDir(pConnPnt->GetMemberAt(0) , vbrace_dir))
					{
						B = pShapeParam->Width;
						ptTmpOrigin += V.Normalize()*(pShapeParam->t1*0.5);
					}

					CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
					double A = (NULL != pParam) ? pParam->A : -1;
					double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
					if((-1 == A) || (-1 == jointL))
					{
						SAFE_DELETE(pGussetPlate);
						return ERROR_BAD_ENVIRONMENT;
					}
					const double cs = cos(alpha) , sn = sin(alpha);
					double totalL = (B*0.5 + SHAPE_LEN_OFFSET + 0.5*A*(cs/sn))*(1/sn) + jointL;
					CIsVect3d cross = vbrace_dir*pGussetPlate->m_norm;
					cross.Normalize();

					tmp[0] = pConnPnt->origin() + vbrace_dir*totalL + cross*A*0.5;
					tmp[1] = pConnPnt->origin() + vbrace_dir*totalL - cross*A*0.5;
				}

				vbrace_dir = pConnPnt->GetMemberAt(2)->Direction();
				if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->start(CSDNFElement::METER)))
				{
					vbrace_dir = -vbrace_dir;
				}
				vbrace_dir.Normalize();
				angle = (beam_dir.AngleTo(vbrace_dir));
				beam_dir = ((angle >= 0.0) && (angle < PI*0.5)) ? beam_dir : -beam_dir;	/// check the angle between beam and hbrace - 2013.06.26 added by humkyung
				//if(!equals(fabs(angle) , PI*0.5))	/// if angle between two vectors isn't 90 degree
				{
					//vbrace_dir = vbrace_dir - pGussetPlate->m_norm.DotProduct(vbrace_dir)*pGussetPlate->m_norm;/// force vbrace_dir to place on plane - 2014.02.11 added by humkyung

					const double alpha = acos(beam_dir.DotProduct(vbrace_dir));

					CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
					double B = 0;
					double dot = vbrace_dir.DotProduct(beam_dir);
					CIsVect3d V = vbrace_dir - beam_dir.Normalize()*dot;
					CIsPoint3d ptTmpOrigin(pConnPnt->origin());
					if(this->IsFlangeDir(pConnPnt->GetMemberAt(0) , vbrace_dir))
					{
						CIsVect3d tmpVec = ptTmpOrigin - pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER);
						dot = fabs(tmpVec.DotProduct(beam_dir));
						CIsPoint3d T(pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER) + beam_dir*dot);
						const double L = fabs(CIsVect3d(T - ptTmpOrigin).DotProduct(V.Normalize()));

						B = (pShapeParam->Height - L)*2;
						ptTmpOrigin += V.Normalize()*(pShapeParam->Height - L);
					}
					else if(this->IsWebDir(pConnPnt->GetMemberAt(0) , vbrace_dir))
					{
						B = pShapeParam->Width;
						ptTmpOrigin += V.Normalize()*(pShapeParam->t1*0.5);
					}

					double A = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(2)->section())->A;
					if(-1 == A) return ERROR_BAD_ENVIRONMENT;
					double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(2)->section());
					if(-1 == jointL) return ERROR_BAD_ENVIRONMENT;
					const double cs = cos(alpha) , sn = sin(alpha);
					double totalL = (B*0.5 + SHAPE_LEN_OFFSET + 0.5*A*(cs/sn))*(1/sn) + jointL;
					
					CIsVect3d cross = vbrace_dir*pGussetPlate->m_norm;
					cross.Normalize();					
					tmp[2] = pConnPnt->origin() + vbrace_dir*totalL + cross*A*0.5;
					tmp[3] = pConnPnt->origin() + vbrace_dir*totalL - cross*A*0.5;
					
					/// align vertices - 2014.02.11 added by humkyung
					int iIndices[2] = {0,};
					double dMin = DBL_MAX;
					for(int i = 0;i < 2;++i)
					{
						for(int j = 0;j < 2;++j)
						{
							const double d = tmp[i].DistanceTo(tmp[2 + j]);
							if(d < dMin)
							{
								dMin = d;
								iIndices[0] = i;
								iIndices[1] = j;
							}
						}
					}
					/// up to here

					pGussetPlate->m_oSectionShapePntList.push_back(tmp[(iIndices[0]+1)%2]);
					pGussetPlate->m_oSectionShapePntList.push_back(tmp[iIndices[0]]);
					pGussetPlate->m_oSectionShapePntList.push_back(tmp[iIndices[1] + 2]);
					pGussetPlate->m_oSectionShapePntList.push_back(tmp[(iIndices[1]+1)%2 + 2]);

					CIsVect3d dir = pGussetPlate->m_oSectionShapePntList[0] - pConnPnt->origin();
					dot = beam_dir.DotProduct(dir);
					pGussetPlate->m_oSectionShapePntList.insert(pGussetPlate->m_oSectionShapePntList.begin() , pConnPnt->origin() + dot*beam_dir);
					dir = pGussetPlate->m_oSectionShapePntList[pGussetPlate->m_oSectionShapePntList.size() - 1] - pConnPnt->origin();
					dot = beam_dir.DotProduct(dir);
					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + dot*beam_dir);
				}
				pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(1)->section());

				oGussetPlateList.push_back(pGussetPlate);
			}
		}
		catch(...)
		{
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	generate gusset plate for VER. BRACE TO VER. BRACE
	@auhor	humkyung
	@date	2013.07.31
*/
int CGussetPlateGenerator::Generate4VBraceToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*> oMemList , CSmartSteelDoc* pDoc)
{
	assert(pDoc && pConnPnt && "pConnPnt or pDoc is NULL");

	if(pDoc && pConnPnt)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
		const CIsPoint3d ptConn(pConnPnt->origin());

		if(3 == oMemList.size())	/// 3 memebrs are connected
		{
			const double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , oMemList[1]->section());
			CIsVect3d lenDir , baseDir , crossDir , secDir;
			lenDir = oMemList[1]->Direction();
			if(ptConn.DistanceTo(oMemList[1]->start(CSDNFElement::METER)) > ptConn.DistanceTo(oMemList[1]->end(CSDNFElement::METER)))
			{
				lenDir.Set(-lenDir.dx() , -lenDir.dy() , -lenDir.dz());
			}
			lenDir.Normalize();
			pGussetPlate->m_norm = oMemList[1]->W();
			pGussetPlate->m_norm.Normalize();
			secDir = pGussetPlate->m_norm*lenDir;

			const double A = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , oMemList[1]->section())->A;
			const double gap = jointL*tan(DEG2RAD(15.0));
			
			pGussetPlate->m_oSectionShapePntList.push_back(ptConn + lenDir*(jointL) + secDir*(A*0.5));
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] - secDir*(A));
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[1] - lenDir*(jointL) - secDir*gap);
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[1] - lenDir*(jointL*2.0));
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[3] + secDir*(A));
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[4] + lenDir*(jointL) + secDir*gap);
			pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(oMemList[1]->section());
			oGussetPlateList.push_back(pGussetPlate);
		}
		else
		{
			double A = 0.0 , B = 0.0 , jointL = 0.0 , totalL = 0.0 , F = 0.0;
			CIsVect3d lenDir , baseDir , crossDir , secDir;
			if((ptConn == oMemList[0]->start(CSDNFElement::METER)) || (ptConn == oMemList[0]->end(CSDNFElement::METER)))
			{
				CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , oMemList[0]->section()/* , _T("A")*/);
				if(NULL == pParam) return ERROR_BAD_ENVIRONMENT;
				A = pParam->A;
				F = A*2.0;
				jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , oMemList[0]->section());
				if((-1 == A) || (-1 == jointL))
				{
					SAFE_DELETE(pGussetPlate);
					return ERROR_BAD_ENVIRONMENT;
				}

				if(ptConn.DistanceTo(oMemList[0]->start(CSDNFElement::METER)) < ptConn.DistanceTo(oMemList[0]->end(CSDNFElement::METER)))
				{
					lenDir = oMemList[0]->end(CSDNFElement::METER) - oMemList[0]->start(CSDNFElement::METER);
				}
				else
				{
					lenDir = oMemList[0]->start(CSDNFElement::METER) - oMemList[0]->end(CSDNFElement::METER);
				}
				lenDir.Normalize();
				baseDir = oMemList[1]->Direction();
				baseDir.Normalize();
				if(lenDir.DotProduct(baseDir) < 0.0) baseDir.Set(-baseDir.dx() , -baseDir.dy() , -baseDir.dz());
				pGussetPlate->m_norm = (baseDir*lenDir).Normalize();
				pGussetPlate->m_norm.Normalize();

				secDir = pGussetPlate->m_norm*baseDir;
			}
			else if((ptConn == oMemList[1]->start(CSDNFElement::METER)) || (ptConn == oMemList[1]->end(CSDNFElement::METER)))
			{
				A = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , oMemList[1]->section())->A;
				F = A*2.0;
				jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , oMemList[1]->section());
				if((-1 == A) || (-1 == jointL))
				{
					SAFE_DELETE(pGussetPlate);
					return ERROR_BAD_ENVIRONMENT;
				}

				if(ptConn.DistanceTo(oMemList[1]->start(CSDNFElement::METER)) < ptConn.DistanceTo(oMemList[1]->end(CSDNFElement::METER)))
				{
					lenDir = oMemList[1]->end(CSDNFElement::METER) - oMemList[1]->start(CSDNFElement::METER);
				}
				else
				{
					lenDir = oMemList[1]->start(CSDNFElement::METER) - oMemList[1]->end(CSDNFElement::METER);
				}
				lenDir.Normalize();
				baseDir = oMemList[0]->Direction();
				baseDir.Normalize();
				if(lenDir.DotProduct(baseDir) < 0.0) baseDir.Set(-baseDir.dx() , -baseDir.dy() , -baseDir.dz());
				pGussetPlate->m_norm = (baseDir*lenDir).Normalize();
				pGussetPlate->m_norm.Normalize();

				secDir = pGussetPlate->m_norm*baseDir;
			}
			else
			{
				SAFE_DELETE(pGussetPlate);
				return ERROR_BAD_ENVIRONMENT;
			}
			crossDir = pGussetPlate->m_norm*lenDir;
			crossDir.Normalize();

			/// connect two lines - 2014.08.26 added by humkyung
			CIsPoint3d tmpP[2][2];
			{
				tmpP[0][0] = ptConn + secDir*(A*0.5 + SHAPE_LEN_OFFSET) + baseDir*(F*0.5);
				tmpP[0][1] = tmpP[0][0] - secDir*(A + SHAPE_LEN_OFFSET);
				tmpP[1][0] = ptConn + lenDir*jointL + crossDir*A*0.5;
				tmpP[1][1] = tmpP[1][0] - crossDir*A;
			}
			double dLen[4] = 
			{
				tmpP[0][0].DistanceTo(tmpP[1][0]),tmpP[0][0].DistanceTo(tmpP[1][1]),
				tmpP[0][1].DistanceTo(tmpP[1][0]),tmpP[0][1].DistanceTo(tmpP[1][1])
			} , dMin = DBL_MAX;

			int iIndex = 4;
			for(int i = 0;i < 4;++i)
			{
				if(dLen[i] < dMin)
				{
					dMin = dLen[i];
					iIndex = i;
				}
			}
			if(0 == iIndex)
			{
				swap(tmpP[0][0] , tmpP[0][1]);
			}
			else if(1 == iIndex)
			{
				swap(tmpP[0][0] , tmpP[0][1]);
				swap(tmpP[1][0] , tmpP[1][1]);
			}
			else if(3 == iIndex)
			{
				swap(tmpP[1][0] , tmpP[1][1]);
			}
			/// up to here

			pGussetPlate->m_oSectionShapePntList.push_back(tmpP[0][0]);
			pGussetPlate->m_oSectionShapePntList.push_back(tmpP[0][1]);
			pGussetPlate->m_oSectionShapePntList.push_back(tmpP[1][0]);
			pGussetPlate->m_oSectionShapePntList.push_back(tmpP[1][1]);
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[1] - baseDir*(F));
			pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] - baseDir*(F));
			pGussetPlate->MakeShapeToConvexHull();

			pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(oMemList[0]->section());

			oGussetPlateList.push_back(pGussetPlate);
		}
	}

	return ERROR_SUCCESS;
}

/**
	@brief	check given connection point is to be column-beam-vbrace
	@auhor	humkyung
	@date	2014.05.21
*/
bool CGussetPlateGenerator::CheckIfColumnBeamVBraceIsPossible(CSteelConnPoint* pConnPnt) const
{
	assert(pConnPnt && "pConnPnt is NULL");
	if(pConnPnt && (3 == pConnPnt->GetMemberSize()))
	{
		const CIsPoint3d ptStart = pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER);
		const CIsPoint3d ptEnd = pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER);

		CAppDocData& docData = CAppDocData::GetInstance();
		if((pConnPnt->origin().DistanceTo(ptStart) < CONN_PNT_TOLER) || (pConnPnt->origin().DistanceTo(ptEnd) < CONN_PNT_TOLER))
		{
			CIsVect3d column_dir = pConnPnt->GetMemberAt(0)->Direction();
			column_dir.Normalize();
			if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER))) column_dir = -column_dir;

			CIsVect3d vbrace_dir = pConnPnt->GetMemberAt(2)->Direction();
			vbrace_dir.Normalize();
			if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->start(CSDNFElement::METER))) vbrace_dir = -vbrace_dir;

			const double dot = column_dir.DotProduct(vbrace_dir);
			return (dot > 0.0);
		}
		return true;
	}

	return false;
}

/**
	@brief	generate gusset plate
	@auhor	humkyung
	@date	2013.06.24
*/
int CGussetPlateGenerator::Generate(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");

	if(pConnPnt && pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		oGussetPlateList.clear();
		switch(pConnPnt->Type())
		{
			case CSteelConnPoint::COLUMN_BEAM_TO_VBRACE:
			{
				try
				{
					if(CheckIfColumnBeamVBraceIsPossible(pConnPnt))///3 == pConnPnt->GetMemberSize())
					{
						CIsVect3d column_dir = pConnPnt->GetMemberAt(0)->Direction();
						column_dir.Normalize();
						CIsVect3d beam_dir   = pConnPnt->GetMemberAt(1)->Direction();
						if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER))) beam_dir = -beam_dir;
						beam_dir.Normalize();
						CIsVect3d vbrace_dir = pConnPnt->GetMemberAt(2)->Direction();
						vbrace_dir.Normalize();
						if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(2)->start(CSDNFElement::METER))) vbrace_dir = -vbrace_dir;
						if(beam_dir.DotProduct(vbrace_dir) < 0.0) beam_dir.Set(-beam_dir.dx() , -beam_dir.dy() , -beam_dir.dz());	/// set beam dir and vbrace dir has same direction.
						const double alpha = acos(beam_dir.DotProduct(vbrace_dir));

						CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(pConnPnt->GetMemberAt(0)->section());
						if(NULL == pShapeParam) return ERROR_BAD_ENVIRONMENT;
						
						const double B = this->GetOffsetFromLinearMember(pConnPnt->GetMemberAt(0) , pConnPnt->GetMemberAt(2)->Type() , pConnPnt->origin() , beam_dir , pDoc);
						CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(2)->section());
						double A = (NULL != pParam) ? pParam->A : -1;
						double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(2)->section());
						if((-1 == A) || (-1 == jointL)) return ERROR_BAD_ENVIRONMENT;
						double totalL = B*0.5*(1/cos(alpha)) + (SHAPE_LEN_OFFSET)*(1/cos(alpha)) + A*sin(alpha) + jointL;
						
						CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
						{
							pGussetPlate->m_norm = pConnPnt->GetMemberAt(2)->W();///beam_dir*vbrace_dir;
							pGussetPlate->m_norm.Normalize();
							CIsVect3d cross = vbrace_dir*pGussetPlate->m_norm;
							cross.Normalize();
							
							pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + beam_dir*(B*0.5));
							pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + (beam_dir*(SHAPE_START_LEN)));
							
							CIsPoint3d tmp[2];
							tmp[0] = pConnPnt->origin() + vbrace_dir*totalL + cross*A*0.5;
							tmp[1] = pConnPnt->origin() + vbrace_dir*totalL - cross*A*0.5;
							double d[2]={0,0};
							d[0] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[0]);
							d[1] = pGussetPlate->m_oSectionShapePntList[1].DistanceTo(tmp[1]);
							if(d[0] < d[1])
							{
								pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
								pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
							}
							else
							{
								pGussetPlate->m_oSectionShapePntList.push_back(tmp[1]);
								pGussetPlate->m_oSectionShapePntList.push_back(tmp[0]);
							}

							CIsVect3d dirV(pGussetPlate->m_oSectionShapePntList[pGussetPlate->m_oSectionShapePntList.size()-1] - pGussetPlate->m_oSectionShapePntList[0]);
							const double len = column_dir.DotProduct(dirV);
							pGussetPlate->m_oSectionShapePntList.push_back(pGussetPlate->m_oSectionShapePntList[0] + column_dir*len);
							
							pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(2)->section());

							oGussetPlateList.push_back(pGussetPlate);
						}
					}
					else
					{
						CSteelConnPoint* pNewConnPnt = pDoc->CreateConnPoint(CSteelConnPoint::BEAM_TO_VBRACE , pConnPnt->origin());
						pNewConnPnt->Add(pConnPnt->GetMemberAt(1));
						pNewConnPnt->Add(pConnPnt->GetMemberAt(2));
						this->Generate(oGussetPlateList , pNewConnPnt , pDoc);
					}
				}
				catch(...)
				{
					int d = 1;
				}
			}
			break;
			case CSteelConnPoint::COLUMN_TO_VBRACE:
			{
				Generate4ColumnToVBrace(oGussetPlateList , pConnPnt , pDoc);
			}
			break;
			case CSteelConnPoint::BEAM_TO_HBRACE:
			{
				vector<CIsPoint3d> pts;
				{
					pts.push_back(pConnPnt->GetMemberAt(0)->start());
					pts.push_back(pConnPnt->GetMemberAt(0)->end());
					pts.push_back((pConnPnt->GetMemberAt(1)->start() + pConnPnt->GetMemberAt(1)->end())*0.5);
				}
				CIsPlane3d tmp;
				CIsPlane3d::CreateOf(tmp , pts);
				CIsPlane3d plane(pConnPnt->origin() , pConnPnt->GetMemberAt(0)->Direction()*tmp.normal());
				
				vector<CSDNFLinearMember*> oPosLinearMemberCol , oNegLinearMemberCol;
				for(int i = 1;i < pConnPnt->GetMemberSize();++i)
				{
					const CIsPoint3d pt = (pConnPnt->GetMemberAt(i)->start() + pConnPnt->GetMemberAt(i)->end())*0.5;
					const CIsPlane3d::SIGN_T sign = plane.WhichSideOf(pt);
					if(CIsPlane3d::POSITIVE == sign)
					{
						oPosLinearMemberCol.push_back(pConnPnt->GetMemberAt(i));
					}
					else if(CIsPlane3d::NEGATIVE == sign)
					{
						oNegLinearMemberCol.push_back(pConnPnt->GetMemberAt(i));
					}
				}

				if(!oPosLinearMemberCol.empty())
				{
					oPosLinearMemberCol.insert(oPosLinearMemberCol.begin() , pConnPnt->GetMemberAt(0));
					Generate4BeamToHBrace(oGussetPlateList , pConnPnt , oPosLinearMemberCol , pDoc);
				}
				if(!oNegLinearMemberCol.empty())
				{
					oNegLinearMemberCol.insert(oNegLinearMemberCol.begin() , pConnPnt->GetMemberAt(0));
					Generate4BeamToHBrace(oGussetPlateList , pConnPnt , oNegLinearMemberCol , pDoc);
				}
			}
			break;
			case CSteelConnPoint::BEAM_TO_VBRACE:
			{
				if(5 == pConnPnt->GetMemberSize())
				{
					CIsVect3d OV = pConnPnt->GetMemberAt(0)->OriVector();
					CIsVect3d U  = CIsVect3d(pConnPnt->GetMemberAt(0)->end() - pConnPnt->GetMemberAt(0)->start()).Normalize();
					CIsVect3d W(U*OV) , V(W*U);
					CIsPlane3d plane(pConnPnt->origin() , V);
					
					vector<CSDNFLinearMember*> oLeftSideMemberList,oRightSideMemberList;
					CIsPoint3d ptMid[4];
					ptMid[0] = (pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER) + pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER))*0.5;
					ptMid[1] = (pConnPnt->GetMemberAt(2)->end(CSDNFElement::METER) + pConnPnt->GetMemberAt(2)->start(CSDNFElement::METER))*0.5;
					ptMid[2] = (pConnPnt->GetMemberAt(3)->end(CSDNFElement::METER) + pConnPnt->GetMemberAt(3)->start(CSDNFElement::METER))*0.5;
					ptMid[3] = (pConnPnt->GetMemberAt(4)->end(CSDNFElement::METER) + pConnPnt->GetMemberAt(4)->start(CSDNFElement::METER))*0.5;
					(CIsPlane3d::POSITIVE == plane.WhichSideOf(ptMid[0])) ? oLeftSideMemberList.push_back(pConnPnt->GetMemberAt(1)) : oRightSideMemberList.push_back(pConnPnt->GetMemberAt(1));
					(CIsPlane3d::POSITIVE == plane.WhichSideOf(ptMid[1])) ? oLeftSideMemberList.push_back(pConnPnt->GetMemberAt(2)) : oRightSideMemberList.push_back(pConnPnt->GetMemberAt(2));
					(CIsPlane3d::POSITIVE == plane.WhichSideOf(ptMid[2])) ? oLeftSideMemberList.push_back(pConnPnt->GetMemberAt(3)) : oRightSideMemberList.push_back(pConnPnt->GetMemberAt(3));
					(CIsPlane3d::POSITIVE == plane.WhichSideOf(ptMid[3])) ? oLeftSideMemberList.push_back(pConnPnt->GetMemberAt(4)) : oRightSideMemberList.push_back(pConnPnt->GetMemberAt(4));

					{
						CSteelConnPoint* pLeftConnPnt(new CSteelConnPoint(pConnPnt->Type() , pConnPnt->origin()));
						if(NULL != pLeftConnPnt)
						{
							pDoc->m_oConnPntList.push_back(pLeftConnPnt);

							pLeftConnPnt->Add(pConnPnt->GetMemberAt(0));
							for(vector<CSDNFLinearMember*>::iterator itr = oLeftSideMemberList.begin();itr != oLeftSideMemberList.end();++itr)
							{
								pLeftConnPnt->Add(*itr);
							}
							Generate4BeamToVBrace(oGussetPlateList , pLeftConnPnt , pDoc);
						}
					}

					{
						CSteelConnPoint* pRightConnPnt(new CSteelConnPoint(pConnPnt->Type() , pConnPnt->origin()));
						if(NULL != pRightConnPnt)
						{
							pDoc->m_oConnPntList.push_back(pRightConnPnt);

							pRightConnPnt->Add(pConnPnt->GetMemberAt(0));
							for(vector<CSDNFLinearMember*>::iterator itr = oRightSideMemberList.begin();itr != oRightSideMemberList.end();++itr)
							{
								pRightConnPnt->Add(*itr);
							}
							Generate4BeamToVBrace(oGussetPlateList , pRightConnPnt , pDoc);
						}
					}
				}
				else
				{
					Generate4BeamToVBrace(oGussetPlateList , pConnPnt , pDoc);
				}
			}
			break;
			case CSteelConnPoint::VBRACE_TO_VBRACE:
			{
				try
				{
					vector<CSDNFLinearMember*> oMemList;
					if(3 == pConnPnt->GetMemberSize())
					{
						/// order member
						int i = 0 , j = 0 , k = 0;
						for(i = 0;i < pConnPnt->GetMemberSize();++i)
						{
							j = (i+1) % 3;
							k = (i+2) % 3;
							CIsVect3d norm = pConnPnt->GetMemberAt(i)->Direction()*pConnPnt->GetMemberAt(j)->Direction();
							if(equals(norm.Length() , 0.0)) break;
						}

						if(3 != i)
						{
							oMemList.push_back(pConnPnt->GetMemberAt(k));
							oMemList.push_back(pConnPnt->GetMemberAt(i));
							oMemList.push_back(pConnPnt->GetMemberAt(j));
							Generate4VBraceToVBrace(oGussetPlateList , pConnPnt , oMemList , pDoc);
						}
						else
						{
							CMainFrame* pFrameWnd = CMainFrame::GetInstance();
							OSTRINGSTREAM_T oss;
							oss << _T("Fail to create gusset plate with ver. brace(") << pConnPnt->GetMemberAt(0)->MemberID() << _T(",");
							oss << pConnPnt->GetMemberAt(1)->MemberID() << _T(",") << pConnPnt->GetMemberAt(2)->MemberID() << _T(")");
							pFrameWnd->SendMessage(DISPLAY_MESSAGE , WPARAM(oss.str().c_str()) , MessageType::MESSAGE_WARNING);
						}
					}
					else
					{
						for(int i = 0;i < pConnPnt->GetMemberSize();++i)
						{
							int j = (i+1) % pConnPnt->GetMemberSize();
							oMemList.clear();
							oMemList.push_back(pConnPnt->GetMemberAt(i));
							oMemList.push_back(pConnPnt->GetMemberAt(j));
							CIsVect3d norm = oMemList[0]->Direction()*oMemList[1]->Direction();
							if(!equals(norm.Length() , 0.0))
							{
								Generate4VBraceToVBrace(oGussetPlateList , pConnPnt , oMemList , pDoc);
							}
						}
					}
				}
				catch(...)
				{
				}
			}
			break;
			default:
			{
				CHorGussetPlateGenerator generator;
				generator.Generate(oGussetPlateList , pConnPnt , pDoc);
			}
			break;
		}

		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	Returns true if the distance between the two points is <br>
	lower or equal to dToler.

	@author	humkyung

	@date	2013.06.26
*/
bool CGussetPlateGenerator::IsEqual(const CIsPoint3d& lhs , const CIsPoint3d& rhs , const double& dToler)
{
	const double dist = lhs.DistanceTo(rhs);
	return (dist <= dToler);
}