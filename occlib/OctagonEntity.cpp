#include "stdafx.h"
#include "OctagonEntity.h"

using namespace OCC;
COctagonEntity::COctagonEntity(void)
{
	m_type = OCC::COctagonEntity::TypeString();

	Reset(NULL);
}

COctagonEntity::~COctagonEntity(void)
{
}

/******************************************************************************
    @author     humkyung
    @date       2014.08.08
    @class      COctagonEntity
    @function   TypeString
    @return     STRING_T
    @brief
******************************************************************************/
STRING_T COctagonEntity::TypeString()
{
	static const STRING_T __type_str__(_T("octagon"));
	
	return __type_str__;
}

int COctagonEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_ptOrigin = gp_Pnt(0.0 , 0.0 , 0.0);
	m_axis = gp_Dir(0.0 , 0.0 , 1.0);
	m_angle = 0.0;
	m_dRadius = 20.0;
	m_dHeight = 5.0;

	return ERROR_SUCCESS;
}

void COctagonEntity::Rotate( const double& angle )
{
	m_angle += angle;
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void COctagonEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	COCCEntity::Rotate( axis , angle );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-10-10
    @class      COctagonEntity
    @function   CreateShape
    @return     int
    @param      BRep_Builder*       pBuilder
    @param      TopoDS_Compound*    pCompound
    @brief
******************************************************************************/
TopoDS_Shape COctagonEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Ax2 ax2 = gp_Ax2(m_ptOrigin , m_axis);
	gp_Dir xdir = ax2.XDirection();
	gp_Dir ydir = ax2.YDirection();

	gp_Pnt pts[8];
	pts[0] = m_ptOrigin;
	pts[0].Translate(-gp_Vec(m_axis) * m_dHeight * 0.5);
	pts[0].Translate(gp_Vec(xdir)*m_dRadius);
	gp_Ax1 ax1(m_ptOrigin , m_axis);
	pts[0].Rotate(ax1 , 22.5*PI180);
	pts[1] = pts[0];
	pts[1].Rotate(ax1 , 45.0*PI180);
	pts[2] = pts[1];
	pts[2].Rotate(ax1 , 45.0*PI180);
	pts[3] = pts[2];
	pts[3].Rotate(ax1 , 45.0*PI180);
	pts[4] = pts[3];
	pts[4].Rotate(ax1 , 45.0*PI180);
	pts[5] = pts[4];
	pts[5].Rotate(ax1 , 45.0*PI180);
	pts[6] = pts[5];
	pts[6].Rotate(ax1 , 45.0*PI180);
	pts[7] = pts[6];
	pts[7].Rotate(ax1 , 45.0*PI180);

	const int iSize = 8;
	for(int i = 0;i < iSize;++i)
	{
		pts[i].Rotate( gp_Ax1(m_ptOrigin , m_axis) , m_angle * PI180 );
	}
	
	BRepBuilderAPI_MakeWire makeStartWire , makeEndWire;
	for(int i = 0;i < iSize;++i)
	{
		gp_Pnt start(pts[i % iSize]);
		gp_Pnt end(pts[(i + 1) % iSize]);
		TopoDS_Edge E = BRepBuilderAPI_MakeEdge(start , end);
		makeStartWire.Add(E);

		start.Translate(gp_Vec(m_axis)*m_dHeight);
		end.Translate(gp_Vec(m_axis)*m_dHeight);
		E = BRepBuilderAPI_MakeEdge(start , end);
		makeEndWire.Add(E);
	}
	
	BRepOffsetAPI_ThruSections generator(Standard_True , Standard_True);
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

/******************************************************************************
    @author     humkyung
    @date       2011-10-10
    @class      COctagonEntity
    @function   Translate
    @return     void
    @param      gp_Vec& V
    @brief
******************************************************************************/
void COctagonEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void COctagonEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
{
	for(vector<CString>::iterator itr = oObject.begin();itr != oObject.end();++itr)
	{
		if(_T("origin") == (*itr))
		{
			m_ptOrigin.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("az") == (*itr))
		{
			m_axis.SetCoord( atof(*(itr + 1)) , atof(*(itr + 2)) , atof(*(itr + 3)) );
			itr+=3;
		}
		else if(_T("angle") == (*itr))
		{
			m_angle = atof(*(itr + 1));
			itr++;
		}
		else if(_T("radius") == (*itr))
		{
			m_dRadius = atof(*(itr + 1));
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

	if(!CreateShape().IsNull()) Redisplay(hContext);
}