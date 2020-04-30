#include "stdafx.h"
#include <assert.h>
#include "CtorEntity.h"

#include "gp_Circ.hxx"

using namespace OCC;

CCtorEntity::CCtorEntity(void)
{
	m_type = _T("ctor");
	m_normal.SetCoord(0,0,1);

	Reset(NULL);
}

CCtorEntity::~CCtorEntity(void)
{
}

int CCtorEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);
	m_ptStart = gp_Pnt(50.0 , 0.0 , 0.0);
	m_dTubeRadius = 5.0;
	m_dSweepAngle = 0.0;

	return ERROR_SUCCESS;
}

void CCtorEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	m_ptStart.Translate(V);
	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CCtorEntity::Rotate( const double& angle )
{
	gp_Vec startVec(m_ptOrigin , m_ptStart);
	m_ptStart.Rotate( gp_Ax1(m_ptOrigin , m_normal) , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CCtorEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	m_ptOrigin.Rotate( axis , angle*PI180 );
	m_ptStart.Rotate( axis , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-26
    @class      CCtorEntity
    @function   CreateShape
    @return     int
    @param      list<Handle(AIS_Shape
    @brief
******************************************************************************/
TopoDS_Shape CCtorEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	try
	{
		gp_Vec startVec(m_ptOrigin , m_ptStart);
		gp_Vec axis = m_normal;
		gp_Vec startAxis = startVec.Crossed(axis);

		gp_Circ c = gp_Circ(gp_Ax2(m_ptStart , gp_Dir(startAxis)) , m_dTubeRadius);
		TopoDS_Edge E = BRepBuilderAPI_MakeEdge(c);
		TopoDS_Wire Wc = BRepBuilderAPI_MakeWire(E);
		TopoDS_Face aStartFace = BRepBuilderAPI_MakeFace(Wc);

		gp_Ax1 axe = gp_Ax1(m_ptOrigin , gp_Dir(axis));
		TopoDS_Shape aShape = BRepPrimAPI_MakeRevol(aStartFace , axe , m_dSweepAngle);
		return aShape;
		
		/*if((NULL != pBuilder) && (NULL != pCompound))
		{
			pBuilder->Add( *pCompound , aShape );
		}
		else
		{
			if(m_oAISShapeList.empty())
			{
				Handle(AIS_Shape) aAISShape = new AIS_Shape(aShape);
				m_oAISShapeList.push_back(aAISShape);
			}
			else
			{
				Handle_AIS_InteractiveContext hContext = m_oAISShapeList[0]->GetContext();

				m_oAISShapeList[0]->Set(aShape);
				if(!hContext.IsNull())
				{
					hContext->Deactivate(m_oAISShapeList[0]);
					hContext->Redisplay(m_oAISShapeList[0] , true , true);
					hContext->Update(m_oAISShapeList[0] , true);
				}
			}
		}
		*/
	}
	catch(...)
	{
	}

	return TopoDS_Shape();
}

void CCtorEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("start") == (*itr))
		{
			m_ptStart.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("sweep") == (*itr))
		{
			m_dSweepAngle = atof(*(itr + 1));
			itr+=3;
		}
		else if(_T("radius") == (*itr))
		{
			m_dTubeRadius = atof(*(itr + 1));
			itr++;
		}
		else if(_T("normal") == (*itr))
		{
			m_normal.SetCoord(atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	if(!CreateShape().IsNull()) Redisplay(hContext);
}

/**
	@brief	return normal vector
	@author	humkyung
	@date	2014.10.16
*/
gp_Vec& CCtorEntity::normal()
{
	return m_normal;
}

/**
	@brief	return sweep angle 
	@author	humkyung
	@date	2014.10.16
*/
double& CCtorEntity::sweepAngle()
{
	return m_dSweepAngle;
}