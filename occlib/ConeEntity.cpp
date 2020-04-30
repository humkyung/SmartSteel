#include "stdafx.h"
#include "occlib.h"
#include <assert.h>
#include "ConeEntity.h"

#include "gp_Circ.hxx"

using namespace OCC;

CConeEntity::CConeEntity(void) : m_dBottomRadius(10.0) , m_dTopRadius(COCCEntity::EPSILON)
{
	m_type = OCC::CConeEntity::TypeString();

	Reset(NULL);
}

CConeEntity::~CConeEntity(void)
{
}

/******************************************************************************
    @brief		return type string of cone
    @author     humkyung
    @date       2014.08.08
    @class      CConeEntity
    @function   TypeString
    @return     STRING_T
******************************************************************************/
STRING_T CConeEntity::TypeString()
{
	static const STRING_T __type_str__(_T("cone"));
	
	return __type_str__;
}

int CConeEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptBottom = gp_Pnt(0 , 0 , -5);
	m_ptTop = gp_Pnt(0 , 0 , 5);
	m_axis = gp_Dir(0.0 , 0.0 , 1.0);

	return ERROR_SUCCESS;
}

void CConeEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	m_ptBottom.Translate(V);
	m_ptTop.Translate(V);
	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CConeEntity::Rotate( const double& angle )
{
	m_ptBottom.Rotate(gp_Ax1(m_ptOrigin , m_axis) , angle*PI180);
	m_ptTop.Rotate(gp_Ax1(m_ptOrigin , m_axis) , angle*PI180);
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CConeEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	m_ptOrigin.Rotate( axis , angle*PI180 );
	m_ptBottom.Rotate( axis , angle*PI180 );
	m_ptTop.Rotate( axis , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CConeEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("bottom") == (*itr))
		{
			m_ptBottom.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("top") == (*itr))
		{
			m_ptTop.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("bradius") == (*itr))
		{
			m_dBottomRadius = atof(*(itr + 1));
			if(0.0 == m_dBottomRadius) m_dBottomRadius = COCCEntity::EPSILON;
			itr++;
		}
		else if(_T("tradius") == (*itr))
		{
			m_dTopRadius = atof(*(itr + 1));
			if(0.0 == m_dTopRadius) m_dTopRadius = COCCEntity::EPSILON;
			itr++;
		}
		else if(_T("az") == (*itr))
		{
			m_axis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}
	m_ptOrigin.SetCoord((m_ptBottom.X() + m_ptTop.X())*0.5 , (m_ptBottom.Y() + m_ptTop.Y())*0.5 , (m_ptBottom.Z() + m_ptTop.Z())*0.5);

	if(!CreateShape().IsNull()) Redisplay(hContext);
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-25
    @class      CConeEntity
    @function   CreateShape
    @return     int
    @brief
******************************************************************************/
TopoDS_Shape CConeEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	try
	{
		gp_Circ c1 = gp_Circ(gp_Ax2(m_ptBottom , m_axis) , (0.0 == m_dBottomRadius) ? CConeEntity::COCCEntity::EPSILON : m_dBottomRadius);
		TopoDS_Edge E1 = BRepBuilderAPI_MakeEdge(c1);
		TopoDS_Wire W1 = BRepBuilderAPI_MakeWire(E1);
		/*TopoDS_Face aBottomFace = BRepBuilderAPI_MakeFace(W1);
		Handle(AIS_Shape) hBottom = new AIS_Shape(aBottomFace);
		oShapeList.push_back(hBottom);*/

		gp_Circ c2 = gp_Circ(gp_Ax2(m_ptTop , m_axis) , (0.0 == m_dTopRadius) ? CConeEntity::COCCEntity::EPSILON : m_dTopRadius);
		TopoDS_Edge E2 = BRepBuilderAPI_MakeEdge(c2);
		TopoDS_Wire W2 = BRepBuilderAPI_MakeWire(E2);
		/*TopoDS_Face aTopFace = BRepBuilderAPI_MakeFace(W2);
		Handle(AIS_Shape) hTop = new AIS_Shape(aTopFace);
		oShapeList.push_back(hTop);*/

		BRepOffsetAPI_ThruSections generator(Standard_True,Standard_True);
		generator.AddWire(W1);
		generator.AddWire(W2);
		try
		{
			generator.Build();
		}
		catch(...)
		{
			int d= 1;
		}
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
				Handle(AIS_Shape) aAISShape = new AIS_ColoredShape(aShape);
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
		int d = 1;
	}
	
	return TopoDS_Shape();
}
