#include "stdafx.h"
#include "ExtruEntity.h"

#include <IsTools.h>

using namespace OCC;
CExtruEntity::CExtruEntity(void)
{
	m_type = _T("extru");

	Reset(NULL);
}

CExtruEntity::~CExtruEntity(void)
{
}

/******************************************************************************
    @brief		return type string of extrusion
    @author     humkyung
    @date       2014.08.08
    @class      CConeEntity
    @function   TypeString
    @return     STRING_T
******************************************************************************/
STRING_T CExtruEntity::TypeString()
{
	static const STRING_T __type_str__(_T("extru"));
	
	return __type_str__;
}

int CExtruEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_oPointList.push_back(gp_Pnt(0.0 , 0.0 , 0.0));
	m_oPointList.push_back(gp_Pnt(30.0 , 0.0 , 0.0));
	m_oPointList.push_back(gp_Pnt(30.0 , 15.0 , 0.0));
	m_oPointList.push_back(gp_Pnt(15.0 , 15.0 , 0.0));
	m_oPointList.push_back(gp_Pnt(15.0 , 30.0 , 0.0));
	m_oPointList.push_back(gp_Pnt(0.0 , 30.0 , 0.0));
	m_axis = gp_Dir(0.0 , 0.0 , 1.0);
	m_dThickness = 5.0;

	return ERROR_SUCCESS;
}

void CExtruEntity::Rotate( const double& angle )
{

}

void CExtruEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	for(vector<gp_Pnt>::iterator itr = m_oPointList.begin();itr != m_oPointList.end();++itr)
	{
		itr->Rotate( axis , angle );
	}
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       ????.??.??
    @class      CExtruEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CExtruEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	if(0.0 == m_dThickness) return TopoDS_Shape();	/// can't create a shape if thickness is zero!!!

	int iEdgeCount = 0;

	BRepBuilderAPI_MakeWire makeStartWire , makeEndWire;
	for(vector<gp_Pnt>::iterator itr = m_oPointList.begin();(itr + 1) != m_oPointList.end();++itr)
	{
		gp_Pnt start(*itr);
		gp_Pnt end(*(itr + 1));
		if(!start.IsEqual(end , COCCEntity::EPSILON))
		{
			TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
			makeStartWire.Add(E);

			start.Translate(gp_Vec(m_axis)*m_dThickness);
			end.Translate(gp_Vec(m_axis)*m_dThickness);
			E = BRepBuilderAPI_MakeEdge(start , end);
			makeEndWire.Add(E);

			++iEdgeCount;
		}
	}
	if(!m_oPointList.back().IsEqual(m_oPointList.front() , COCCEntity::EPSILON))
	{
		gp_Pnt start(m_oPointList.back());
		gp_Pnt end(m_oPointList.front());
		TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
		makeStartWire.Add(E);

		start.Translate(gp_Vec(m_axis)*m_dThickness);
		end.Translate(gp_Vec(m_axis)*m_dThickness);
		E = BRepBuilderAPI_MakeEdge(start , end);
		makeEndWire.Add(E);

		++iEdgeCount;
	}
	
	if(iEdgeCount > 0)
	{
		BRepOffsetAPI_ThruSections generator(((true == m_bSolid) ? Standard_True : Standard_False) , Standard_True);
		generator.AddWire(makeStartWire.Wire());
		generator.AddWire(makeEndWire.Wire());
		generator.Build();
		TopoDS_Shape aShape = generator.Shape();
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
				aAISShape->SetTransparency( m_dTransparency );
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
		}*/
	}

	return TopoDS_Shape();
}

void CExtruEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);	
	for(vector<gp_Pnt>::iterator itr = m_oPointList.begin();itr != m_oPointList.end();++itr)
	{
		itr->Translate(V);
	}

	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CExtruEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{

}