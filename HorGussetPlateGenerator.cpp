#include "StdAfx.h"
#include <assert.h>
#include <IsPlane3d.h>
#include "HorGussetPlateGenerator.h"
#include "AppDocData.h"
#include "GussetPlateGenerator.h"

CHorGussetPlateGenerator::CHorGussetPlateGenerator(void)
{
}

CHorGussetPlateGenerator::~CHorGussetPlateGenerator(void)
{
}

/**
	@brief	find a base linear member
	@author	humkyung
	@date	2014.03.25
*/
CSDNFLinearMember* CHorGussetPlateGenerator::FindBaseMember(CSteelConnPoint* pConnPnt)
{
	assert(pConnPnt && "pConnPnt is NULL");
	CSDNFLinearMember* res = NULL;

	if(pConnPnt)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		for(int i = 0;i < pConnPnt->GetMemberSize();++i)
		{
			if(!CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(i)->start(CSDNFElement::METER) , CONN_PNT_TOLER) && !CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(i)->end(CSDNFElement::METER) , CONN_PNT_TOLER))
			{
				res = pConnPnt->GetMemberAt(i);
				break;
			}
		}
	}

	return res;
}

/**
	@brief	generate horizontal gusset plate
	@author	humkyung
	@date	2014.03.25
*/
int CHorGussetPlateGenerator::Generate(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");

	if(pConnPnt && pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		double A = 0.0 , B = 0.0 , jointL = 0.0 , totalL = 0.0 , F = 0.0;
		switch(pConnPnt->Type())
		{
			case CSteelConnPoint::HBRACE_TO_HBRACE:
			{
				if(2 == pConnPnt->GetMemberSize())
				{
					CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
					
					CIsVect3d lenDir , baseDir , crossDir;
					if(CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER) , CONN_PNT_TOLER) || CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER) , CONN_PNT_TOLER))
					{
						CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(0)->section());
						A = (NULL != pParam) ? pParam->A : -1.0;
						F = A*2.0;
						jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(0)->section());
						if((-1 == A) || (-1 == jointL))
						{
							SAFE_DELETE(pGussetPlate);
							return ERROR_BAD_ENVIRONMENT;
						}

						if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER)))
						{
							lenDir = pConnPnt->GetMemberAt(0)->Direction();
						}
						else
						{
							lenDir = pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER) - pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER);
						}
						lenDir.Normalize();
						baseDir = (pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER));
						baseDir.Normalize();
						pGussetPlate->m_norm = baseDir*lenDir;
						pGussetPlate->m_norm.Normalize();
					}
					else if(CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER) , CONN_PNT_TOLER) || CGussetPlateGenerator::IsEqual(pConnPnt->origin() , pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER) , CONN_PNT_TOLER))
					{
						CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
						A = (NULL != pParam) ? pParam->A : -1;
						F = A*2.0;
						jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pConnPnt->GetMemberAt(1)->section());
						if((-1 == A) || (-1 == jointL))
						{
							SAFE_DELETE(pGussetPlate);
							return ERROR_BAD_ENVIRONMENT;
						}

						if(pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER)))
						{
							lenDir = pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER);
						}
						else
						{
							lenDir = pConnPnt->GetMemberAt(1)->start(CSDNFElement::METER) - pConnPnt->GetMemberAt(1)->end(CSDNFElement::METER);
						}
						lenDir.Normalize();
						baseDir = (pConnPnt->GetMemberAt(0)->end(CSDNFElement::METER) - pConnPnt->GetMemberAt(0)->start(CSDNFElement::METER));
						baseDir.Normalize();
						pGussetPlate->m_norm = baseDir*lenDir;
						pGussetPlate->m_norm.Normalize();
					}
					else
					{
						SAFE_DELETE(pGussetPlate);
						return ERROR_BAD_ENVIRONMENT;
					}
					crossDir = pGussetPlate->m_norm*lenDir;
					crossDir.Normalize();

					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() + baseDir*(F*0.5));
					CIsPoint3d tmpP[2];
					tmpP[0] = pConnPnt->origin() + lenDir*jointL + crossDir*A*0.5;
					tmpP[1] = tmpP[0] - crossDir*A;
					if(pGussetPlate->m_oSectionShapePntList[0].DistanceTo(tmpP[0]) < pGussetPlate->m_oSectionShapePntList[0].DistanceTo(tmpP[1]))
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmpP[0]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmpP[1]);
					}
					else
					{
						pGussetPlate->m_oSectionShapePntList.push_back(tmpP[1]);
						pGussetPlate->m_oSectionShapePntList.push_back(tmpP[0]);
					}
					pGussetPlate->m_oSectionShapePntList.push_back(pConnPnt->origin() - baseDir*(F*0.5));

					pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pConnPnt->GetMemberAt(0)->section());

					oGussetPlateList.push_back(pGussetPlate);
				}
				/// 2014.03.25 added by humkyung
				else if(3 == pConnPnt->GetMemberSize())
				{
					/// generate a gusset plate for 3 hor. braces
					CSDNFLinearMember* pBaseMem = FindBaseMember(pConnPnt);
					CSDNFLinearMember* pSideMem[2] = {NULL,NULL};
					int iSideMem = 0;
					for(int i = 0;i < pConnPnt->GetMemberSize();++i)
					{
						if(pBaseMem != pConnPnt->GetMemberAt(i))
						{
							pSideMem[iSideMem++] = pConnPnt->GetMemberAt(i);
						}
					}
					if((NULL != pBaseMem) && (NULL != pSideMem[0]) && (NULL != pSideMem[1]))
					{
						CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pSideMem[0]->section());
						A = (NULL != pParam) ? pParam->A : -1;
						F = A*2.0;
						jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pSideMem[0]->section());
						if((-1 == A) || (-1 == jointL)) return ERROR_BAD_ENVIRONMENT;

						CIsPoint3d tmp[6];
						CIsVect3d crossDir , lenDir;
						CGussetPlate* pGussetPlate = new CGussetPlate(pConnPnt);
						{
							pGussetPlate->m_norm = pBaseMem->Direction()*pSideMem[0]->Direction();
							pGussetPlate->m_norm.Normalize();
							
							crossDir = pBaseMem->Direction();///(pGussetPlate->m_norm*pSideMem[0]->Direction()).Normalize();
							if(pConnPnt->origin().DistanceTo(pSideMem[0]->start(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pSideMem[0]->end(CSDNFElement::METER)))
							{
								lenDir = pSideMem[0]->end(CSDNFElement::METER) - pSideMem[0]->start(CSDNFElement::METER);
							}
							else
							{
								lenDir = pSideMem[0]->start(CSDNFElement::METER) - pSideMem[0]->end(CSDNFElement::METER);
							}
							lenDir.Normalize();

							tmp[0] = pConnPnt->origin() + lenDir.Normalize()*jointL + crossDir*A*0.5;
							tmp[1] = pConnPnt->origin() + lenDir.Normalize()*jointL - crossDir*A*0.5;
						}

						pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , pSideMem[1]->section());
						A = (NULL != pParam) ? pParam->A : -1;
						F = A*2.0;
						jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , pSideMem[1]->section());
						if((-1 == A) || (-1 == jointL)) return ERROR_BAD_ENVIRONMENT;

						if(pConnPnt->origin().DistanceTo(pSideMem[1]->start(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(pSideMem[1]->end(CSDNFElement::METER)))
						{
							lenDir = pSideMem[1]->end(CSDNFElement::METER) - pSideMem[1]->start(CSDNFElement::METER);
						}
						else
						{
							lenDir = pSideMem[1]->start(CSDNFElement::METER) - pSideMem[1]->end(CSDNFElement::METER);
						}
						lenDir.Normalize();
						tmp[3] = pConnPnt->origin() + lenDir.Normalize()*jointL - crossDir*A*0.5;
						tmp[4] = pConnPnt->origin() + lenDir.Normalize()*jointL + crossDir*A*0.5;
						
						tmp[2] = pConnPnt->origin() - pBaseMem->Direction()*F*0.5;
						tmp[5] = pConnPnt->origin() + pBaseMem->Direction()*F*0.5;
						/*const double dist1 = tmp[1].DistanceTo(tmp[2]);
						const double dist2 = tmp[1].DistanceTo(tmp[5]);
						if(dist1 > dist2) swap(tmp[2] , tmp[5]);*/
						for(int i = 0;i < SIZE_OF_ARRAY(tmp);++i)
						{
							pGussetPlate->m_oSectionShapePntList.push_back(tmp[i]);
						}

						pGussetPlate->m_dThickness = pDoc->GetPlateThicknessOf(pBaseMem->section());
						oGussetPlateList.push_back(pGussetPlate);
					}
				}
				/// up to here
			}
		}
	}

	return ERROR_SUCCESS;
}

/**
	@brief	get outer point list of gussetplate
	@author	humkyung
	@date	2014.02.11
*/
int CHorGussetPlateGenerator::GetOuterPntListOfPlate(vector<CIsPoint3d>& oShapePntList , CIsVect3d& norm , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*>& oMemList , CSmartSteelDoc* pDoc)
{
	assert(pConnPnt && pDoc && "pConnPnt or pDoc is NULL");
	oShapePntList.clear();

	if(pConnPnt && pDoc && (oMemList.size() >= 2))
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		vector<CIsPoint3d> pts;
		{
			pts.push_back(oMemList[0]->start());
			pts.push_back(oMemList[0]->end());
			pts.push_back((oMemList[1]->start() + oMemList[1]->end())*0.5);
			pts[0].SetZ(pConnPnt->origin().z());
			pts[1].SetZ(pConnPnt->origin().z());
			pts[2].SetZ(pConnPnt->origin().z());
		}
		CIsPlane3d plane;
		CIsPlane3d::CreateOf(plane , pts);
		norm = plane.normal();

		CIsVect3d beam_dir = oMemList[0]->Direction();
		beam_dir.Normalize();
		for(int i = 1;i < int(oMemList.size());++i)
		{
			CIsVect3d hbrace_dir = oMemList[i]->Direction();
			if(pConnPnt->origin().DistanceTo(oMemList[i]->end(CSDNFElement::METER)) < pConnPnt->origin().DistanceTo(oMemList[i]->start(CSDNFElement::METER)))
			{
				hbrace_dir = -hbrace_dir;
			}
			hbrace_dir.Normalize();
			const double angle = (beam_dir.AngleTo(hbrace_dir));
			beam_dir = (beam_dir.DotProduct(hbrace_dir) < 0) ? -beam_dir : beam_dir;	/// check the angle between beam and hbrace - 2013.06.26 added by humkyung
			
			/// force to locate gusset plate over linear member - 2013.07.28 added by humkyung
			if(norm.dz() < 0.0) norm.Set(-norm.dx() , -norm.dy() , -norm.dz());
			CIsVect3d base_dir = norm*beam_dir;
			base_dir.Normalize();
			if(base_dir.DotProduct(hbrace_dir) < 0.0) base_dir.Set(-base_dir.dx() , -base_dir.dy() , -base_dir.dz());

			const double alpha = acos(base_dir.DotProduct(hbrace_dir));

			CSteelSectionBuilder::ShapeParam* pShapeParam = pDoc->GetShapeParamOf(oMemList[0]->section());
			const double offset = CGussetPlateGenerator::GetOffsetFromLinearMember(oMemList[0] , oMemList[i]->Type() , pConnPnt->origin() , base_dir , pDoc);
			CGussetPlate::Param* pParam = pDoc->GetSteelValue(CSDNFLinearMember::VBRACE , oMemList[i]->section());
			const double A = (NULL != pParam) ? pParam->A : -1;
			double jointL = pDoc->GetSteelJointLength(CSDNFLinearMember::VBRACE , oMemList[i]->section());
			if((-1 == A) || (-1 == jointL)) return ERROR_BAD_ENVIRONMENT;

			const double B = (CGussetPlateGenerator::IsFlangeDir(oMemList[0] , hbrace_dir)) ? pShapeParam->Height : pShapeParam->Width;
			const double totalL = B*0.5*(1/cos(alpha)) + (SHAPE_LEN_OFFSET)*(1/cos(alpha)) + A*sin(alpha) + jointL;
			CIsVect3d cross = hbrace_dir*norm;
			cross.Normalize();

			/*oShapePntList.push_back(pConnPnt->origin() + base_dir*(offset*0.5) - beam_dir*(SHAPE_START_LEN));
			oShapePntList.push_back(oShapePntList[0] + (base_dir*(B*0.5 + SHAPE_START_LEN)));*/

			CIsPoint3d tmp[4];
			tmp[0] = pConnPnt->origin() + hbrace_dir*totalL + cross*A*0.5;
			tmp[1] = pConnPnt->origin() + hbrace_dir*totalL - cross*A*0.5;
			if(!oShapePntList.empty())
			{
				tmp[3] = oShapePntList[oShapePntList.size() - 1];
				oShapePntList.pop_back();
				tmp[2] = oShapePntList[oShapePntList.size() - 1];
				oShapePntList.pop_back();

				/// align vertices - 2014.02.11 added by humkyung
				int iIndices[2] = {0,};
				double dMin = DBL_MAX;
				for(int i = 0;i < 2;++i)
				{
					for(int j = 0;j < 2;++j)
					{
						const double d = tmp[2 + i].DistanceTo(tmp[j]);
						if(d < dMin)
						{
							dMin = d;
							iIndices[0] = i;
							iIndices[1] = j;
						}
					}
				}
				/// up to here
				oShapePntList.push_back(tmp[(iIndices[0]+1)%2 + 2]);
				oShapePntList.push_back(tmp[iIndices[0] + 2]);
				oShapePntList.push_back(tmp[iIndices[1]]);
				oShapePntList.push_back(tmp[(iIndices[1]+1)%2]);
			}
			else
			{
				oShapePntList.push_back(tmp[0]);
				oShapePntList.push_back(tmp[1]);
			}
		}
		return ERROR_SUCCESS;
	}

	return ERROR_BAD_ENVIRONMENT;
}