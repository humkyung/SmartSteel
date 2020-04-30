#include "StdAfx.h"
#include <assert.h>
#include <Geom_Line.hxx>
#include <GeomAPI_ExtremaCurveCurve.hxx>
#include <gp_Dir.hxx>
#include <IsPlane3d.h>
#include "ConnPntGenerator.h"
#include "SmartSteelDoc.h"
#include "OCCShapeBuilder.h"
#include "AppDocData.h"

#include <OCCEntFactory.h>
#include <ComplexShapeEntity.h>

const double CConnPntGenerator::TOLER(0.00001);

CConnPntGenerator::CConnPntGenerator(void)
{
}

CConnPntGenerator::~CConnPntGenerator(void)
{
}

/******************************************************************************
    @brief		check duplicate origin when type is less then given type
	@author     humkyung
    @date       2013-06-19
    @class      CSmartSteelDoc
    @function   CheckDuplicate
    @return     bool
******************************************************************************/
CSteelConnPoint* CConnPntGenerator::CheckDuplicate(CSmartSteelDoc* pDoc , const int& iType , const CIsPoint3d& pt , list<CSteelConnPoint*>* pConnPntList , CSDNFLinearMember* pMem1 , CSDNFLinearMember* pMem2)
{
	list<CSteelConnPoint*>* pTempConnPntList = (NULL != pConnPntList) ? pConnPntList : &(pDoc->m_oConnPntList);
	for(list<CSteelConnPoint*>::iterator itr = pTempConnPntList->begin();itr != pTempConnPntList->end();++itr)
	{
		if((*itr)->Type() <= iType)
		{
			if((*itr)->origin() == pt)
			{
				if((NULL == pMem1) && (NULL == pMem2)) return (*itr);
				if((NULL != pMem1) && (*itr)->HasMember(pMem1)) return (*itr);
				if((NULL != pMem2) && (*itr)->HasMember(pMem2)) return (*itr);
				return NULL;
			}
		}
	}

	return NULL;
}

/******************************************************************************
    @brief		check duplicate
	@author     humkyung
    @date       2013-07-14
    @class      CSmartSteelDoc
    @function   CheckDuplicate
    @return     bool
******************************************************************************/
CSteelConnPoint* CConnPntGenerator::CheckDuplicate(CSmartSteelDoc* pDoc , const int& iType , const vector<CSDNFLinearMember*>& oMemList)
{
	for(list<CSteelConnPoint*>::iterator itr = pDoc->m_oConnPntList.begin();itr != pDoc->m_oConnPntList.end();++itr)
	{
		if((*itr)->Type() < iType)
		{
			bool bFound = true;
			for(vector<CSDNFLinearMember*>::const_iterator jtr = oMemList.begin();jtr != oMemList.end();++jtr)
			{
				if(!(*itr)->HasMember(*jtr))
				{
					bFound = false;
					break;
				}
			}
			if(true == bFound) return (*itr);
		}
	}

	return NULL;
}

/******************************************************************************
    @brief		get volume of linear member
	@author     humkyung
    @date       2013-07-10
    @class      CSmartSteelDoc
    @function   GetVolumeOf
    @return     CIsVolume
******************************************************************************/
CIsVolume CConnPntGenerator::GetVolumeOf(CSmartSteelDoc* pDoc , CSDNFLinearMember* pLinearMem)
{
	assert(pDoc && pLinearMem && "pDoc or pLinear Member is NULL");

	CIsVolume res;
	if(pDoc && pLinearMem)
	{
		CSteelSectionBuilder oSteelShapeBuilder;
		oSteelShapeBuilder.Build(pLinearMem , pDoc);
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

/******************************************************************************
    @brief		extend given line if not blocked by
	@author     humkyung
    @date       2013-07-15
    @class      CSmartSteelDoc
    @function   ExtendLine
    @return     int
******************************************************************************/
int CConnPntGenerator::ExtendLine(CSmartSteelDoc* pDoc , CIsLine3d& line , const CSDNFLinearMember::ElmType& iBlockLinearType)
{
	assert(pDoc && "pDoc is NULL");
	if(pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator where = pDoc->m_oSDNFElmMap.find(iBlockLinearType);
		if(where != pDoc->m_oSDNFElmMap.end())
		{
			const CIsPoint3d start = line.start();
			const CIsPoint3d end = line.end();
			bool bBlocked[2] = {false , false};
			for(list<CSDNFLinearMember*>::iterator itr = where->second->begin();itr != where->second->end();++itr)
			{
				if((true == bBlocked[0]) && (true == bBlocked[1])) break;
				CIsVolume vol = GetVolumeOf(pDoc , *itr);
				if(vol.Contains(start))
				{
					bBlocked[0] = true;
				}
				else if(vol.Contains(end))
				{
					bBlocked[1] = true;
				}
			}
			
			CIsVect3d dir = line.Direction().Normalize();
			line.Set(start - ((false == bBlocked[0]) ? dir*EXTEND_OFFSET : CIsPoint3d(0.0,0.0,0.0)) , end + ((false == bBlocked[1]) ? dir*EXTEND_OFFSET : CIsPoint3d(0.0,0.0,0.0)));

			return ERROR_SUCCESS;
		}
		else
		{
			const CIsPoint3d start = line.start();
			const CIsPoint3d end = line.end();

			CIsVect3d dir = line.Direction().Normalize();
			line.Set(start - dir*EXTEND_OFFSET , end + dir*EXTEND_OFFSET);
		}
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	post process
	@author	humkyung
	@date	2013.07.13
*/
int CConnPntGenerator::PostProcess(CSmartSteelDoc* pDoc)
{
	assert(pDoc && "pDoc is NULL");
	if(pDoc)
	{
		CAppDocData& docData = CAppDocData::GetInstance();

		map<CSDNFLinearMember* , list<CSteelConnPoint*>* > oBeamAssociateMap;
		for(list<CSteelConnPoint*>::iterator itr = pDoc->m_oConnPntList.begin();itr != pDoc->m_oConnPntList.end();++itr)
		{
			if((CSteelConnPoint::BEAM_TO_VBRACE == (*itr)->Type()) && (2 == (*itr)->GetMemberSize()))
			{
				map<CSDNFLinearMember* , list<CSteelConnPoint*>* >::iterator where = oBeamAssociateMap.find((*itr)->GetMemberAt(0));
				if(where == oBeamAssociateMap.end())
				{
					list<CSteelConnPoint*>* pConnPntList = new list<CSteelConnPoint*>;
					pConnPntList->push_back(*itr);
					oBeamAssociateMap.insert(make_pair((*itr)->GetMemberAt(0) , pConnPntList));
				}
				else
				{
					where->second->push_back(*itr);
				}
			}
		}

		for(map<CSDNFLinearMember* , list<CSteelConnPoint*>* >::iterator itr = oBeamAssociateMap.begin();itr != oBeamAssociateMap.end();++itr)
		{
			CIsVolume vol = GetVolumeOf(pDoc , itr->first);
			for(list<CSteelConnPoint*>::iterator jtr = itr->second->begin();jtr != itr->second->end();++jtr)
			{
				CIsLine3d vbrace_line1((*jtr)->GetMemberAt(1)->start(CSDNFElement::METER) , (*jtr)->GetMemberAt(1)->end(CSDNFElement::METER));
				CIsVect3d dir = vbrace_line1.Direction().Normalize();
				vbrace_line1.Set(vbrace_line1.start() - dir*EXTEND_OFFSET , vbrace_line1.end() + dir*EXTEND_OFFSET);
				for(list<CSteelConnPoint*>::iterator ktr = jtr;ktr != itr->second->end();)
				{
					if(jtr == ktr)
					{
						++ktr;
						continue;
					}
					
					CIsLine3d vbrace_line2((*ktr)->GetMemberAt(1)->start(CSDNFElement::METER) , (*ktr)->GetMemberAt(1)->end(CSDNFElement::METER));
					dir = vbrace_line2.Direction().Normalize();
					vbrace_line2.Set(vbrace_line2.start() - dir*EXTEND_OFFSET , vbrace_line2.end() + dir*EXTEND_OFFSET);

					CIsPoint3d intsect;
					if(INTERSECT == vbrace_line1.IntersectWith(intsect , vbrace_line2))
					{
						if(vol.Contains(intsect))
						{
							(*jtr)->origin() = intsect;
							(*jtr)->Add((*ktr)->GetMemberAt(1));

							pDoc->m_oConnPntList.remove(*ktr);
							SAFE_DELETE(*ktr);

							ktr = itr->second->erase(ktr);
							continue;
						}
					}
					++ktr;
				}
			}
		}

		/// free allocated memory
		for(map<CSDNFLinearMember* , list<CSteelConnPoint*>* >::iterator itr = oBeamAssociateMap.begin();itr != oBeamAssociateMap.end();++itr)
		{
			SAFE_DELETE(itr->second);
		}
		oBeamAssociateMap.clear();
	}

	return ERROR_INVALID_PARAMETER;
}

/**
	@brief	get intersection point of two lines
	@author	humkyung
	@date	2014.08.27
*/
bool CConnPntGenerator::IntersectWith(CIsPoint3d& intsect , double& dist , const CIsLine3d& lhs , const CIsLine3d& rhs , const double& dToler) const
{
	Handle(Geom_Line) aLhsLine = new Geom_Line(gp_Pnt(lhs.start().x() , lhs.start().y() , lhs.start().z()) , 
		gp_Dir(lhs.end().x() - lhs.start().x() , lhs.end().y() - lhs.start().y() , lhs.end().z() - lhs.start().z()));
	Handle(Geom_Line) aRhsLine = new Geom_Line(gp_Pnt(rhs.start().x() , rhs.start().y() , rhs.start().z()) , 
		gp_Dir(rhs.end().x() - rhs.start().x() , rhs.end().y() - rhs.start().y() , rhs.end().z() - rhs.start().z()));
	GeomAPI_ExtremaCurveCurve aIntersection(aLhsLine , aRhsLine);
	if((aIntersection.NbExtrema() > 0) && (aIntersection.LowerDistance() < dToler))
	{
		dist = aIntersection.LowerDistance();
		gp_Pnt aIntersect[2];
		aIntersection.Points(1 , aIntersect[0] , aIntersect[1]);
		
		CIsVolume vol[2];
		CIsVect3d vec(lhs.end() - lhs.start());
		vec = vec.Normalize();
		vol[0].Add(lhs.start() - vec*dToler);vol[0].Add(lhs.end() + vec*dToler);
		vec.Set(rhs.end().x() - rhs.start().x() , rhs.end().y() - rhs.start().y() , rhs.end().z() - rhs.start().z());
		vec = vec.Normalize();
		vol[1].Add(rhs.start() - vec*dToler);vol[1].Add(rhs.end() + vec*dToler);
		if( vol[0].Contains(CIsPoint3d(aIntersect[0].X() , aIntersect[0].Y() , aIntersect[0].Z())) && 
			vol[1].Contains(CIsPoint3d(aIntersect[1].X() , aIntersect[1].Y() , aIntersect[1].Z())))
		{
			intsect.Set(aIntersect[1].X() , aIntersect[1].Y() , aIntersect[1].Z());
			return true;
		}
	}

	return false;
}

/**
	@brief	generate connection point
	@author	humkyung
	@date	2013.07.12
*/
int CConnPntGenerator::Generate(CSmartSteelDoc* pDoc)
{
	assert(pDoc && "pDoc is NULL");
	if(pDoc)
	{
		OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
		CAppDocData& docData = CAppDocData::GetInstance();

		///(1) COLUMN & BEAM TO VER. BRACE
		map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator column_where = pDoc->m_oSDNFElmMap.find(CSDNFLinearMember::COLUMN);
		map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator beam_where = pDoc->m_oSDNFElmMap.find(CSDNFLinearMember::BEAM);
		map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator vbrace_where = pDoc->m_oSDNFElmMap.find(CSDNFLinearMember::VBRACE);
		map<CSDNFLinearMember::ElmType , list<CSDNFLinearMember*>* >::iterator hbrace_where = pDoc->m_oSDNFElmMap.find(CSDNFLinearMember::HBRACE);
		if((column_where != pDoc->m_oSDNFElmMap.end()) && (beam_where != pDoc->m_oSDNFElmMap.end()) && (vbrace_where != pDoc->m_oSDNFElmMap.end()))
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = column_where->second->begin();itr != column_where->second->end();++itr)
			{
				CIsLine3d column_line((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				CIsVect3d column_dir(column_line.end() - column_line.start());
				column_dir.Normalize();
				for(list<CSDNFLinearMember*>::iterator jtr = beam_where->second->begin();jtr != beam_where->second->end();++jtr)
				{
					CIsLine3d beam_line((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					if(INTERSECT == column_line.IntersectWith(intsect , beam_line))
					{
						for(list<CSDNFLinearMember*>::iterator ktr = vbrace_where->second->begin();ktr != vbrace_where->second->end();++ktr)
						{
							if((intsect == (*ktr)->start(CSDNFElement::METER)) || (intsect == (*ktr)->end(CSDNFElement::METER)))
							{
								CIsLine3d vbrace_line((*ktr)->start(CSDNFElement::METER) , (*ktr)->end(CSDNFElement::METER));
								CIsVect3d beam_dir   = (intsect == beam_line.end()) ? -beam_line.Direction() : beam_line.Direction();
								CIsVect3d vbrace_dir = (intsect == vbrace_line.end()) ? -vbrace_line.Direction() : vbrace_line.Direction();
								double rad = beam_dir.AngleTo(vbrace_dir);
								rad = (beam_dir.DotProduct(vbrace_dir) < 0.0) ? (rad - PI*0.5) : rad;
								const double dot = (beam_dir*vbrace_dir).DotProduct(column_dir);
								if(equals(dot , 0.0) && rad < PI*0.5)
								{
									CSteelConnPoint* pConnPnt = new CSteelConnPoint(CSteelConnPoint::COLUMN_BEAM_TO_VBRACE , intsect);
									if(pConnPnt)
									{
										pConnPnt->Add(*itr);
										pConnPnt->Add(*jtr);
										pConnPnt->Add(*ktr);
										pDoc->m_oConnPntList.push_back(pConnPnt);
									}
								}
							}
						}
					}
				}
			}
		}

		///(2) COLUMN TO VER. BRACE
		if((column_where != pDoc->m_oSDNFElmMap.end()) && (vbrace_where != pDoc->m_oSDNFElmMap.end()))
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = column_where->second->begin();itr != column_where->second->end();++itr)
			{
				CIsVolume vol(this->GetVolumeOf(pDoc , *itr));
				CIsLine3d column_line((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				for(list<CSDNFLinearMember*>::iterator jtr = vbrace_where->second->begin();jtr != vbrace_where->second->end();++jtr)
				{
					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

					CIsLine3d vbrace_line((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					if(vol.Contains(vbrace_line.start()) || vol.Contains(vbrace_line.end()))
					{
						intsect = vol.Contains(vbrace_line.start()) ? vbrace_line.start() : vbrace_line.end();
						CSteelConnPoint* pConnPnt = CheckDuplicate(pDoc , CSteelConnPoint::COLUMN_TO_VBRACE , intsect);
						if((NULL == pConnPnt) || (CSteelConnPoint::COLUMN_TO_VBRACE == pConnPnt->Type()))
						{
							CSteelConnPoint* pConnPnt = new CSteelConnPoint(CSteelConnPoint::COLUMN_TO_VBRACE , intsect);
							if(pConnPnt)
							{
								pConnPnt->Add(*itr);
								pConnPnt->Add(*jtr);
								pDoc->m_oConnPntList.push_back(pConnPnt);
							}
						}
					}
				}
			}
		}

		///(3) BEAM TO VER. BRACE
		if((beam_where != pDoc->m_oSDNFElmMap.end()) && (vbrace_where != pDoc->m_oSDNFElmMap.end()))
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = beam_where->second->begin();itr != beam_where->second->end();++itr)
			{
				CIsVolume vol(this->GetVolumeOf(pDoc , *itr));
				CIsLine3d beam_line((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				for(list<CSDNFLinearMember*>::iterator jtr = vbrace_where->second->begin();jtr != vbrace_where->second->end();++jtr)
				{
					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

					CIsLine3d vbrace_line((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					if(INTERSECT == beam_line.IntersectWith(intsect , vbrace_line))
					{
						if(false == CheckDuplicate(pDoc , CSteelConnPoint::BEAM_TO_VBRACE , intsect))
						{
							CSteelConnPoint* pConnPnt = new CSteelConnPoint(CSteelConnPoint::BEAM_TO_VBRACE , intsect);
							if(pConnPnt)
							{
								pConnPnt->Add(*itr);
								pConnPnt->Add(*jtr);
								pDoc->m_oConnPntList.push_back(pConnPnt);
							}
						}
					}
					else if(vol.Contains(vbrace_line.start()) || vol.Contains(vbrace_line.end()))
					{
						CIsLine3d copy_beam_line(beam_line) , copy_vbrace_line(vbrace_line); 
						ExtendLine(pDoc , copy_beam_line , CSDNFLinearMember::COLUMN);
						ExtendLine(pDoc , copy_vbrace_line , CSDNFLinearMember::COLUMN);
						Handle(Geom_Line) aBeamLine = new Geom_Line(gp_Pnt(copy_beam_line.start().x() , copy_beam_line.start().y() , copy_beam_line.start().z()) , 
							gp_Dir(copy_beam_line.end().x() - copy_beam_line.start().x() , copy_beam_line.end().y() - copy_beam_line.start().y() , copy_beam_line.end().z() - copy_beam_line.start().z()));
						Handle(Geom_Line) aVBraceLine = new Geom_Line(gp_Pnt(copy_vbrace_line.start().x() , copy_vbrace_line.start().y() , copy_vbrace_line.start().z()) , 
							gp_Dir(copy_vbrace_line.end().x() - copy_vbrace_line.start().x() , copy_vbrace_line.end().y() - copy_vbrace_line.start().y() , copy_vbrace_line.end().z() - copy_vbrace_line.start().z()));
						GeomAPI_ExtremaCurveCurve aIntersection( aBeamLine, aVBraceLine );
						if((aIntersection.NbExtrema() > 0) && (aIntersection.LowerDistance() < TOLER))
						{
							gp_Pnt aIntersect[2];
							aIntersection.Points(1 , aIntersect[0] , aIntersect[1]);
							CIsPoint3d intsect(aIntersect[0].X() , aIntersect[0].Y() , aIntersect[0].Z());
							if(vbrace_line.start().DistanceTo(intsect) < vbrace_line.end().DistanceTo(intsect))
							{
								intsect = vbrace_line.start();
							}
							else
							{
								intsect = vbrace_line.end();
							}
							

							vector<CSDNFLinearMember*> oMemList;
							oMemList.push_back(*itr);
							oMemList.push_back(*jtr);
							if(false == CheckDuplicate(pDoc , CSteelConnPoint::BEAM_TO_VBRACE , intsect))
							{
								CSteelConnPoint* pConnPnt = new CSteelConnPoint(CSteelConnPoint::BEAM_TO_VBRACE , intsect);
								if(pConnPnt)
								{
									pConnPnt->Add(*itr);
									pConnPnt->Add(*jtr);
									pDoc->m_oConnPntList.push_back(pConnPnt);
								}
							}
						}
					}
				}
			}
		}

		///(4) BEAM TO HOR. BRACE
		if((beam_where != pDoc->m_oSDNFElmMap.end()) && (hbrace_where != pDoc->m_oSDNFElmMap.end()))
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = beam_where->second->begin();itr != beam_where->second->end();++itr)
			{
				CIsLine3d beam_line((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				CIsVolume vol = GetVolumeOf(pDoc , *itr);
				for(list<CSDNFLinearMember*>::iterator jtr = hbrace_where->second->begin();jtr != hbrace_where->second->end();++jtr)
				{
					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;
					
					pShape = pDoc->GetShapeParamOf((*itr)->section());
					if(NULL == pShape) continue;
					const double dToler = pShape->Height;

					double dist = 0;
					CIsLine3d hbrace_line((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					if(IntersectWith(intsect , dist , beam_line , hbrace_line , dToler))
					{
						if(hbrace_line.start().DistanceTo(intsect) < hbrace_line.end().DistanceTo(intsect))
						{
							dist = hbrace_line.start().DistanceTo(intsect);
							intsect = hbrace_line.start();
						}
						else
						{
							dist = hbrace_line.end().DistanceTo(intsect);
							intsect = hbrace_line.end();
						}

						if(dist < dToler)
						{
							CSteelConnPoint* pConnPnt  = CheckDuplicate(pDoc , CSteelConnPoint::BEAM_TO_HBRACE , intsect);
							if(NULL == pConnPnt)
							{
								pConnPnt = new CSteelConnPoint(CSteelConnPoint::BEAM_TO_HBRACE , intsect);
								if(pConnPnt)
								{
									pConnPnt->m_dist = dist;
									pConnPnt->Add(*itr);
									pConnPnt->Add(*jtr);
									pDoc->m_oConnPntList.push_back(pConnPnt);
								}
							}
							else
							{
								///refer #48(choose the beam which has shortest distance from conn. point if several beams are connected)
								if((pConnPnt->GetMemberAt(0) != *itr) && (dist < pConnPnt->m_dist))
								{
									pConnPnt->SetMemberAt(0 , *itr);
								}
								/// up to here
								pConnPnt->Add(*jtr);
							}
						}
					}
				}
			}
		}

		///(5) VER. BRACE TO VER. BRACE
		list<CSteelConnPoint*> oVertConnPntList;
		if(vbrace_where != pDoc->m_oSDNFElmMap.end())
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = vbrace_where->second->begin();itr != vbrace_where->second->end();++itr)
			{
				CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*itr)->section());
				if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

				CIsLine3d vbrace_line1((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				list<CSDNFLinearMember*>::iterator jtr = itr;
				for(++jtr;jtr != vbrace_where->second->end();++jtr)
				{
					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

					CIsLine3d vbrace_line2((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					if(INTERSECT == vbrace_line1.IntersectWith(intsect , vbrace_line2))
					{
						CSteelConnPoint* pConnPnt = CheckDuplicate(pDoc , CSteelConnPoint::VBRACE_TO_VBRACE , intsect , &oVertConnPntList);
						if(NULL == pConnPnt)
						{
							pConnPnt = new CSteelConnPoint(CSteelConnPoint::VBRACE_TO_VBRACE , intsect );
							if(pConnPnt)
							{
								pConnPnt->Add(*itr);
								pConnPnt->Add(*jtr);
								oVertConnPntList.push_back(pConnPnt);
							}
						}
						else
						{
							pConnPnt->Add(*itr);
							pConnPnt->Add(*jtr);
						}
					}
				}
			}
		}

		///(6) HOR. BRACE TO HOR. BRACE
		if(hbrace_where != pDoc->m_oSDNFElmMap.end())
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = hbrace_where->second->begin();itr != hbrace_where->second->end();++itr)
			{
				CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*itr)->section());
				if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

				CIsLine3d hbrace_line1((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				hbrace_line1[0] = hbrace_line1[0] - hbrace_line1.Direction().Normalize()*CONN_PNT_TOLER;
				hbrace_line1[1] = hbrace_line1[1] + hbrace_line1.Direction().Normalize()*CONN_PNT_TOLER;
				list<CSDNFLinearMember*>::iterator jtr = itr;
				for(++jtr;jtr != hbrace_where->second->end();++jtr)
				{
					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if((NULL != pShape) && (!docData.m_oPlateCfg.generate_for_web_type_brace && (_T("WFB") == pShape->Shape))) continue;

					CIsLine3d hbrace_line2((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));
					hbrace_line2[0] = hbrace_line2[0] - hbrace_line2.Direction().Normalize()*CONN_PNT_TOLER;
					hbrace_line2[1] = hbrace_line2[1] + hbrace_line2.Direction().Normalize()*CONN_PNT_TOLER;
					if(INTERSECT == hbrace_line1.IntersectWith(intsect , hbrace_line2))
					{
						bool bConnectedToBeam = false;
						if(beam_where != pDoc->m_oSDNFElmMap.end())
						{
							for(list<CSDNFLinearMember*>::iterator ktr = beam_where->second->begin();ktr != beam_where->second->end();++ktr)
							{
								CIsVolume vol(GetVolumeOf(pDoc , *ktr));
								if(vol.Contains(intsect))
								{
									bConnectedToBeam = true;
									break;
								}
							}
						}
						
						if((false == bConnectedToBeam))
						{
							CSteelConnPoint* pConnPnt = CheckDuplicate(pDoc , CSteelConnPoint::HBRACE_TO_HBRACE , intsect , NULL , *itr , *jtr);
							if(NULL == pConnPnt)
							{
								pConnPnt = new CSteelConnPoint(CSteelConnPoint::HBRACE_TO_HBRACE , intsect);
								if(pConnPnt)
								{
									pConnPnt->Add(*itr);
									pConnPnt->Add(*jtr);
									pDoc->m_oConnPntList.push_back(pConnPnt);
								}
							}
							else
							{
								pConnPnt->Add(*itr);
								pConnPnt->Add(*jtr);
							}
						}
					}
				}
			}
		}

		///(7) END PLATE
		if((column_where != pDoc->m_oSDNFElmMap.end()) && (beam_where != pDoc->m_oSDNFElmMap.end()))
		{
			CIsPoint3d intsect;
			for(list<CSDNFLinearMember*>::iterator itr = column_where->second->begin();itr != column_where->second->end();++itr)
			{
				CIsLine3d column_line((*itr)->start(CSDNFElement::METER) , (*itr)->end(CSDNFElement::METER));
				CIsVect3d column_dir(column_line.end() - column_line.start());
				column_dir.Normalize();
				for(list<CSDNFLinearMember*>::iterator jtr = beam_where->second->begin();jtr != beam_where->second->end();++jtr)
				{
					CIsLine3d beam_line((*jtr)->start(CSDNFElement::METER) , (*jtr)->end(CSDNFElement::METER));

					CSteelSectionBuilder::ShapeParam* pShape = pDoc->GetShapeParamOf((*jtr)->section());
					if(NULL == pShape) continue;
					const double dToler = pShape->Height;
					double dist = 0;
					if(IntersectWith(intsect , dist , beam_line , column_line , dToler))
					{
						if(beam_line.start().DistanceTo(intsect) < beam_line.end().DistanceTo(intsect))
						{
							dist = beam_line.start().DistanceTo(intsect);
							intsect = beam_line.start();
						}
						else
						{
							dist = beam_line.end().DistanceTo(intsect);
							intsect = beam_line.end();
						}

						if(dist < dToler)
						{
							CSteelConnPoint* pConnPnt = new CSteelConnPoint(CSteelConnPoint::COLUMN_TO_BEAM , intsect);
							if(pConnPnt)
							{
								pConnPnt->Add(*itr);
								pConnPnt->Add(*jtr);
								pDoc->m_oConnPntList.push_back(pConnPnt);
							}
						}
					}
				}
			}
		}

		/// check vert to vert connection point is located inside beam. if so, input beam to connection point as member
		for(list<CSteelConnPoint*>::iterator itr = oVertConnPntList.begin();itr != oVertConnPntList.end();++itr)
		{
			bool bDeleted = false;
			if(beam_where != pDoc->m_oSDNFElmMap.end())
			{
				for(list<CSDNFLinearMember*>::iterator jtr = beam_where->second->begin();jtr != beam_where->second->end();++jtr)
				{
					CIsVect3d dir = (*jtr)->Direction();
					dir.Normalize();

					CIsVolume volume = GetVolumeOf(pDoc , *jtr);
					CIsVolume tmp;
					tmp.Add((*itr)->origin());
					/// origin point is located in volume of beam and two ver. brace and beam are on one plane.
					const double dot = ((*itr)->GetMemberAt(0)->Direction()*(*itr)->GetMemberAt(1)->Direction()).DotProduct(dir);
					if(volume.CollideWith(tmp) && (dot <= TOLER))
					{
						CIsVect3d norm = dir*(dir*((*itr)->GetMemberAt(0)->Direction()));
						CIsPlane3d plane((*itr)->origin() , norm.Normalize());

						CIsPlane3d::SIGN_T sign[2]={CIsPlane3d::POSITIVE , CIsPlane3d::POSITIVE};
						CIsPoint3d pts[2];
						pts[0] = ((*itr)->GetMemberAt(0)->start(CSDNFElement::METER) + (*itr)->GetMemberAt(0)->end(CSDNFElement::METER))*0.5;
						pts[1] = ((*itr)->GetMemberAt(1)->start(CSDNFElement::METER) + (*itr)->GetMemberAt(1)->end(CSDNFElement::METER))*0.5;
						sign[0] = plane.WhichSideOf(pts[0]);
						sign[1] = plane.WhichSideOf(pts[1]);
						if(sign[0] == sign[1])
						{
							/// ticket #6 - 2013.07.07 added by humkyung
							for(list<CSteelConnPoint*>::iterator ktr = pDoc->m_oConnPntList.begin();ktr != pDoc->m_oConnPntList.end();/*++ktr*/)
							{
								if((CSteelConnPoint::BEAM_TO_VBRACE == (*ktr)->Type()) && (2 == (*ktr)->GetMemberSize()))
								{
									if((*jtr == (*ktr)->GetMemberAt(0)) && (((*itr)->GetMemberAt(0) == (*ktr)->GetMemberAt(1)) || ((*itr)->GetMemberAt(1) == (*ktr)->GetMemberAt(1))))
									{
										SAFE_DELETE(*ktr);
										ktr = pDoc->m_oConnPntList.erase(ktr);
										continue;
									}
								}
								++ktr;
							}
							/// up to here

							(*itr)->Type() = CSteelConnPoint::BEAM_TO_VBRACE;
							(*itr)->Insert(0 , *jtr);

							break;
						}
						else
						{
							SAFE_DELETE(*itr);
							bDeleted = true;
							break;
						}
					}
				}
			}
			if(true == bDeleted) continue;

			if(column_where != pDoc->m_oSDNFElmMap.end())
			{
				for(list<CSDNFLinearMember*>::iterator jtr = column_where->second->begin();jtr != column_where->second->end();++jtr)
				{
					CIsVolume volume = GetVolumeOf(pDoc , *jtr);
					CIsVolume tmp;
					tmp.Add((*itr)->origin());
					if(!volume.IsEmpty() && volume.CollideWith(tmp))
					{
						SAFE_DELETE(*itr);
						bDeleted = true;
						break;
					}
				}
			}
			if(true == bDeleted) continue;

			CSteelConnPoint* pConnPnt = CheckDuplicate(pDoc , CSteelConnPoint::VBRACE_TO_VBRACE , (*itr)->origin());
			if(NULL == pConnPnt)
			{
				pDoc->m_oConnPntList.push_back(*itr);
			}
			else
			{
				SAFE_DELETE(*itr);
			}
		}
		
		/// related ticket #11 - 2013.07.13 added by humkyung
		PostProcess(pDoc);
		
		return ERROR_SUCCESS;
	}

	return ERROR_INVALID_PARAMETER;
}