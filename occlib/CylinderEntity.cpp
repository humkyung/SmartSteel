#include "stdafx.h"
#include "CylinderEntity.h"

#include "gp_Circ.hxx"

using namespace OCC;

CCylinderEntity::CCylinderEntity(void)
{
	m_type = _T("cylinder");

	Reset(NULL);
}

CCylinderEntity::~CCylinderEntity(void)
{
}

/******************************************************************************
    @brief		return type string of cylinder
    @author     humkyung
    @date       2014.08.08
    @class      CConeEntity
    @function   TypeString
    @return     STRING_T
******************************************************************************/
STRING_T CCylinderEntity::TypeString()
{
	static const STRING_T __type_str__(_T("cylinder"));
	
	return __type_str__;
}

int CCylinderEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptStart = gp_Pnt(0.0 , 0.0 , -5.0);
	m_ptEnd   = gp_Pnt(0.0 , 0.0 , 5.0);
	m_dRadius = 5.0;

	return ERROR_SUCCESS;
}

void CCylinderEntity::Rotate( const double& angle )
{
}

void CCylinderEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	COCCEntity::Rotate( axis , angle );

	m_ptStart.Rotate( axis , angle*PI180 );
	m_ptEnd.Rotate( axis , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CCylinderEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
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
		else if(_T("end") == (*itr))
		{
			m_ptEnd.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("radius") == (*itr))
		{
			m_dRadius = atof(*(itr + 1));
			itr++;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	if(!CreateShape().IsNull()) Redisplay(hContext);
}

void CCylinderEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	m_ptStart.Translate(V);
	m_ptEnd.Translate(V);
	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-10-10
    @class      CCylinderEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CCylinderEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Vec aVec(m_ptStart , m_ptEnd);
	const Standard_Real dLength = aVec.Magnitude();
	gp_Dir axis(aVec);
	if(dLength > 0.0)
	{
		gp_Circ c1 = gp_Circ(gp_Ax2(m_ptStart , axis) , m_dRadius);
		TopoDS_Edge aEdge1 = BRepBuilderAPI_MakeEdge(c1);
		TopoDS_Wire aWire1 = BRepBuilderAPI_MakeWire(aEdge1);

		gp_Circ c2 = gp_Circ(gp_Ax2(m_ptEnd , axis) , m_dRadius);
		TopoDS_Edge aEdge2 = BRepBuilderAPI_MakeEdge(c2);
		TopoDS_Wire aWire2 = BRepBuilderAPI_MakeWire(aEdge2);

		BRepOffsetAPI_ThruSections generator(((true == m_bSolid) ? Standard_True : Standard_False) , Standard_True);
		generator.AddWire(aWire1);
		generator.AddWire(aWire2);
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

		return ERROR_SUCCESS;*/
	}
	return TopoDS_Shape();
	//return ERROR_BAD_ENVIRONMENT;
}
