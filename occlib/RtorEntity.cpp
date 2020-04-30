#include "stdafx.h"
#include "RtorEntity.h"

using namespace OCC;
CRtorEntity::CRtorEntity(void)
{
	m_type = OCC::CRtorEntity::TypeString();

	Reset(NULL);
}

CRtorEntity::~CRtorEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2014.08.08
    @class      CRtorEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CRtorEntity::TypeString()
{
	static const STRING_T __type_str__(_T("rtor"));
	
	return __type_str__;
}

int CRtorEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_dWidth = 10.0;
	m_dHeight = 10.0;
	m_ptP1 = gp_Pnt(50.0 , 0.0 , 0.0);
	m_ptP2 = gp_Pnt(0.0 , 50.0 , 0.0);

	return ERROR_SUCCESS;
}

void CRtorEntity::Rotate( const double& angle )
{
	gp_Vec startVec(m_ptOrigin , m_ptP1);
	gp_Vec endVec(m_ptOrigin , m_ptP2);
	gp_Vec axis = startVec.Crossed(endVec);

	m_ptP1.Rotate( gp_Ax1(m_ptOrigin , axis) , angle*PI180 );
	m_ptP2.Rotate( gp_Ax1(m_ptOrigin , axis) , angle*PI180 );
	/*if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}*/
}

void CRtorEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	m_ptOrigin.Rotate( axis , angle*PI180 );
	m_ptP1.Rotate( axis , angle*PI180 );
	m_ptP2.Rotate( axis , angle*PI180 );
	/*if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}*/
}

/******************************************************************************
    @author     humkyung
    @date       2011-08-26
    @class      CRtorEntity
    @function   CreateShape
    @return     int
    @param      list<Handle(AIS_Shape
    @brief
******************************************************************************/
TopoDS_Shape CRtorEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Vec startVec(m_ptOrigin , m_ptP1);
	gp_Vec endVec(m_ptOrigin , m_ptP2);
	gp_Vec axis = startVec.Crossed(endVec);
	gp_Vec startAxis = startVec.Crossed(axis);

	gp_Pnt vertices[4];
	vertices[0] = m_ptP1;
	vertices[0].Translate(-gp_Vec(gp_Dir(startVec))*0.5*m_dWidth);
	vertices[0].Translate(-gp_Vec(gp_Dir(axis))*0.5*m_dHeight);
	vertices[1] = vertices[0];
	vertices[1].Translate(gp_Vec(gp_Dir(startVec))*m_dWidth);
	vertices[2] = vertices[1];
	vertices[2].Translate(gp_Vec(gp_Dir(axis))*m_dHeight);
	vertices[3] = vertices[2];
	vertices[3].Translate(-gp_Vec(gp_Dir(startVec))*m_dWidth);

	TopoDS_Edge E1 = BRepBuilderAPI_MakeEdge(vertices[0] , vertices[1]);
	TopoDS_Edge E2 = BRepBuilderAPI_MakeEdge(vertices[1] , vertices[2]);
	TopoDS_Edge E3 = BRepBuilderAPI_MakeEdge(vertices[2] , vertices[3]);
	TopoDS_Edge E4 = BRepBuilderAPI_MakeEdge(vertices[3] , vertices[0]);
	TopoDS_Wire W = BRepBuilderAPI_MakeWire(E1 , E2 , E3 , E4);
	TopoDS_Face aStartFace = BRepBuilderAPI_MakeFace(W);
			
	Standard_Real angle = startVec.Angle(endVec);

	gp_Ax1 axe = gp_Ax1(m_ptOrigin , gp_Dir(axis));
	TopoDS_Shape aShape = BRepPrimAPI_MakeRevol(aStartFace , axe , angle);
	return aShape;
	
	/*if((NULL != pBuilder) && (NULL != pCompound))
	{
		pBuilder->Add( *pCompound , aShape );
	}
	else
	{
		if(m_oAISShapeList.empty())
		{
			Handle(AIS_Shape) aAISShape = new AIS_Shape( aShape);
			m_oAISShapeList.push_back(aAISShape);
		}
		else
		{
			Handle_AIS_InteractiveContext hContext = m_oAISShapeList[0]->GetContext();

			m_oAISShapeList[0]->Set(aShape);
			if(!hContext.IsNull())
			{
				///hContext->Deactivate(m_oAISShapeList[0]);
				hContext->Redisplay(m_oAISShapeList[0] , true , true);
				hContext->Update(m_oAISShapeList[0] , true);
				///hContext->CurrentViewer()->Redraw();
			}
		}
	}

	return ERROR_SUCCESS;*/
}

void CRtorEntity::Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V)
{
	m_ptOrigin.Translate(V);
	m_ptP1.Translate(V);
	m_ptP2.Translate(V);
	
	///if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CRtorEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("pt1") == (*itr))
		{
			m_ptP1.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("pt2") == (*itr))
		{
			m_ptP2.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("width") == (*itr))
		{
			m_dWidth = atof(*(itr + 1));
			itr++;
		}
		else if(_T("height") == (*itr))
		{
			m_dHeight = atof(*(itr + 1));
			itr++;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	///if(ERROR_SUCCESS == CreateShape()) Display(hContext , 0);
}