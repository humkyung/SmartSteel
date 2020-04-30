#include "StdAfx.h"
#include "GussetPlate.h"

#include <OCCEntFactory.h>
#include <ComplexShapeEntity.h>
#include "OCCShapeBuilder.h"
#include "AppDocData.h"

#include <fstream>
using namespace std;

CGussetPlate::CGussetPlate(CSteelConnPoint* pConnPnt) : CSteelPlate(pConnPnt)
{
	m_sType = _T("Gusset Plate");
}

CGussetPlate::~CGussetPlate(void)
{
	try
	{
		for(vector<OCC::CComplexShapeEntity* >::iterator itr = m_oShapeEntList.begin();itr != m_oShapeEntList.end();++itr)
		{
			SAFE_DELETE(*itr);
		}
		m_oShapeEntList.clear();
	}
	catch(...)
	{
	}
}

/**
	@brief	make gusset plate shape to convex hull
	@author	humkyung
	@date	2014.03.27
*/
void CGussetPlate::MakeShapeToConvexHull()
{
	CAppDocData& docData = CAppDocData::GetInstance();

	CIsVect3d oBaseVec(m_oSectionShapePntList.back().x() - (*m_oSectionShapePntList.begin()).x() , m_oSectionShapePntList.back().y() - (*m_oSectionShapePntList.begin()).y(), m_oSectionShapePntList.back().z() - (*m_oSectionShapePntList.begin()).z());
	oBaseVec = oBaseVec.Normalize();

	bool bLoop = false;
	do
	{
		bLoop = false;
		for(vector<CIsPoint3d>::iterator itr = (m_oSectionShapePntList.begin() + 1);itr != (m_oSectionShapePntList.end() - 1);)
		{
			CIsVect3d oTargetVec(*itr - (*m_oSectionShapePntList.begin()));
			const double dDot = oBaseVec.DotProduct(oTargetVec);
			if(dDot < -CONN_PNT_TOLER)
			{
				m_oSectionShapePntList[0] += oBaseVec*dDot;
				m_oSectionShapePntList.erase((m_oSectionShapePntList.begin() + 1) , itr);
				bLoop = true;
				break;
			}
			else
			{
				++itr;
			}
		}
	}
	while(bLoop);
}

/**
	@brief	check if has given shape
	@author	humkyung
	@date	2013.07.22
*/
bool CGussetPlate::HasShape(const TopoDS_Shape& selectedShape)
{
	for(vector<OCC::CComplexShapeEntity* >::iterator itr = m_oShapeEntList.begin();itr != m_oShapeEntList.end();++itr)
	{
		if((*itr)->m_hShape == selectedShape)
		{
			return true;
		}
	}

	return false;
}

/**
	@brief	display gusset plate
	@author	humkyung
	@date	2013.06.25
*/
int CGussetPlate::Display(Handle(AIS_InteractiveContext) hContext)
{
	OCC::COCCEntFactory& factory = OCC::COCCEntFactory::Instance();
	
	CAppDocData& docData = CAppDocData::GetInstance();
	STRING_T color(_T("255,0,255"));
	try
	{
		CAppDocData::ColorQuad colorQuad = docData.GetColorOf(docData.m_oPlateCfg.gusset_plate_display_color_);
		
		OSTRINGSTREAM_T oss;
		oss << colorQuad.red << _T(",") << colorQuad.green << _T(",") << colorQuad.blue;
		color = oss.str();
	}
	catch(std::invalid_argument)
	{
	}
	
	try
	{
		if(!m_oSectionShapePntList.empty())
		{
			OCC::CComplexShapeEntity* pComplexShape = (OCC::CComplexShapeEntity*)factory.GetEntity( OCC::CComplexShapeEntity::TypeString() );
			if(pComplexShape)
			{
				COCCShapeBuilder builder;

				pComplexShape->SetColor(color.c_str());
				pComplexShape->m_hShape = builder.Shape(m_oSectionShapePntList , m_norm , m_dThickness);
				if(!pComplexShape->m_hShape.IsNull())
				{
					pComplexShape->Display(hContext , AIS_Shaded);
					m_oShapeEntList.push_back(pComplexShape);

					return ERROR_SUCCESS;
				}
			}
		}
	}
	catch(...)
	{
	}

	return ERROR_BAD_ENVIRONMENT;
}

/**
	@brief	show or hide AIS_Shape
	@author	humkyung
	@date	2013.07.27
*/
int CGussetPlate::Show(Handle(AIS_InteractiveContext) hContext , const bool& bShow)
{
	for(vector<OCC::CComplexShapeEntity*>::iterator itr = m_oShapeEntList.begin();itr != m_oShapeEntList.end();++itr)
	{
		(*itr)->Show(hContext , bShow);
	}

	return ERROR_SUCCESS;
}

/**
	@brief	write gusset plate data to m3d file
	@author	humkyung
	@date	2013.06.25
*/
int CGussetPlate::Write(OFSTREAM_T& ofile , const double dUnitScale)
{
	if(ofile.is_open() && !m_oSectionShapePntList.empty() && (CSteelPlate::ALIVE == status()))
	{
		ofile << _T("{") << std::endl;
		{
			ofile << _T("NAME=GUSSET PLATE") << std::endl;
			ofile << _T("SITE=SITE") << std::endl;
			ofile << _T("ZONE=ZONE") << std::endl;
			ofile << _T("POSITION=") << _T("E ") << m_oSectionShapePntList[0].x()*dUnitScale << _T(" N ") << m_oSectionShapePntList[0].y()*dUnitScale << _T(" U ") << m_oSectionShapePntList[0].z()*dUnitScale << std::endl;
			ofile << _T("ORIGIN=") << m_norm.dx() << _T(" ") << m_norm.dy() << _T(" ") << m_norm.dz() << std::endl;
			ofile << _T("TYPE=FWP") << std::endl;
			ofile << _T("[") << std::endl;
			{
				ofile << _T("SubName=GUSSET PLATE") << std::endl;
				ofile << _T("Purpose=Footing") << std::endl;
				ofile << _T("EXTR=") << m_dThickness*dUnitScale << _T(",");
				ofile << _T("E ") << m_oSectionShapePntList[0].x()*dUnitScale << _T(" N ") << m_oSectionShapePntList[0].y()*dUnitScale << _T(" U ") << m_oSectionShapePntList[0].z()*dUnitScale << _T(",");
				ofile << m_norm.dx() << _T(" ") << m_norm.dy() << _T(" ") << m_norm.dz() << _T(",GPLATE") << std::endl;
				ofile << _T("(") << std::endl;
				{
					for(vector<CIsPoint3d>::iterator itr = m_oSectionShapePntList.begin();itr != m_oSectionShapePntList.end();++itr)
					{
						ofile << _T("VERT=") << _T("E ") << itr->x()*dUnitScale << _T(" N ") << itr->y()*dUnitScale << _T(" U ") << itr->z()*dUnitScale << std::endl;
					}
				}
				ofile << _T(")") << std::endl;
			}
			ofile << _T("]") << std::endl;
		}
		ofile << _T("}") << std::endl;

		return ERROR_SUCCESS;
	}

	return ERROR_BAD_ENVIRONMENT;
}