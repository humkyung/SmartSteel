#include "stdafx.h"
#include "PyramidEntity.h"

using namespace OCC;
CPyramidEntity::CPyramidEntity(void)
{
	m_type = OCC::CPyramidEntity::TypeString();

	Reset(NULL);
}

CPyramidEntity::~CPyramidEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2014.08.08
    @class      CPyramidEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T CPyramidEntity::TypeString()
{
	static const STRING_T __type_str__(_T("pyramid"));
	
	return __type_str__;
}

int CPyramidEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptBottom = gp_Pnt(0.0 , 0.0 , -50.0);
	m_ptTop = gp_Pnt(-10.0 , 0.0 , 50.0);
	m_ptOrigin.SetCoord( (m_ptBottom.X() + m_ptTop.X())*0.5 , (m_ptBottom.Y() + m_ptTop.Y())*0.5 , (m_ptBottom.Z() + m_ptTop.Z())*0.5 );

	m_dWidth1 = 50.0, m_dHeight1 = 30.0;
	m_dWidth2 = 20.0, m_dHeight2 = 20.0;
	m_xAxis = gp_Dir(1.0 , 0.0 , 0.0);
	m_yAxis = gp_Dir(0.0 , 1.0 , 0.0);

	return ERROR_SUCCESS;
}

void CPyramidEntity::Rotate( const double& angle )
{
	gp_Dir az = m_xAxis.Crossed( m_yAxis );
	m_ptBottom.Rotate( gp_Ax1(m_ptOrigin , az) , angle*PI180 );
	m_ptTop.Rotate( gp_Ax1(m_ptOrigin , az) , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CPyramidEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	COCCEntity::Rotate( axis , angle );
	m_ptBottom.Rotate( axis , angle*PI180 );
	m_ptTop.Rotate( axis , angle*PI180 );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-30
    @class      CPyramidEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape CPyramidEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Pnt vertices[4];
	vertices[0] = m_ptBottom;
	vertices[0].Translate(-(gp_Vec(m_xAxis)*0.5*m_dWidth1 + gp_Vec(m_yAxis)*0.5*m_dHeight1));
	vertices[1] = vertices[0];
	vertices[1].Translate(gp_Vec(m_xAxis)*m_dWidth1);
	vertices[2] = vertices[1];
	vertices[2].Translate(gp_Vec(gp_Dir(m_yAxis))*m_dHeight1);
	vertices[3] = vertices[2];
	vertices[3].Translate(-gp_Vec(m_xAxis)*m_dWidth1);

	TopoDS_Edge E1 = BRepBuilderAPI_MakeEdge(vertices[0] , vertices[1]);
	TopoDS_Edge E2 = BRepBuilderAPI_MakeEdge(vertices[1] , vertices[2]);
	TopoDS_Edge E3 = BRepBuilderAPI_MakeEdge(vertices[2] , vertices[3]);
	TopoDS_Edge E4 = BRepBuilderAPI_MakeEdge(vertices[3] , vertices[0]);
	TopoDS_Wire aBottomWire = BRepBuilderAPI_MakeWire(E1 , E2 , E3 , E4);

	vertices[0] = m_ptTop;
	vertices[0].Translate(-(gp_Vec(m_xAxis)*0.5*m_dWidth2 + gp_Vec(m_yAxis)*0.5*m_dHeight2));
	vertices[1] = vertices[0];
	vertices[1].Translate(gp_Vec(m_xAxis)*m_dWidth2);
	vertices[2] = vertices[1];
	vertices[2].Translate(gp_Vec(gp_Dir(m_yAxis))*m_dHeight2);
	vertices[3] = vertices[2];
	vertices[3].Translate(-gp_Vec(m_xAxis)*m_dWidth2);

	E1 = BRepBuilderAPI_MakeEdge(vertices[0] , vertices[1]);
	E2 = BRepBuilderAPI_MakeEdge(vertices[1] , vertices[2]);
	E3 = BRepBuilderAPI_MakeEdge(vertices[2] , vertices[3]);
	E4 = BRepBuilderAPI_MakeEdge(vertices[3] , vertices[0]);
	TopoDS_Wire aTopWire = BRepBuilderAPI_MakeWire(E1 , E2 , E3 , E4);

	BRepOffsetAPI_ThruSections generator(Standard_True , Standard_True);
	generator.AddWire(aBottomWire);
	generator.AddWire(aTopWire);
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
			m_oAISShapeList[0]->Set(aShape);
		}
	}

	return ERROR_SUCCESS;*/
}

void CPyramidEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);	
	m_ptBottom.Translate(V);	
	m_ptTop.Translate(V);	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CPyramidEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
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
		else if(_T("width1") == (*itr))
		{
			m_dWidth1 = atof(*(itr + 1));
			itr++;
		}
		else if(_T("height1") == (*itr))
		{
			m_dHeight1 = atof(*(itr + 1));
			itr++;
		}
		else if(_T("width2") == (*itr))
		{
			m_dWidth2 = atof(*(itr + 1));
			itr++;
		}
		else if(_T("height2") == (*itr))
		{
			m_dHeight2 = atof(*(itr + 1));
			itr++;
		}
		else if(_T("ax") == (*itr))
		{
			m_xAxis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("ay") == (*itr))
		{
			m_yAxis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("color") == (*itr))
		{
			m_sColor = *(itr + 1);
			itr++;
		}
	}

	if(!CreateShape().IsNull()) Redisplay(hContext);
}