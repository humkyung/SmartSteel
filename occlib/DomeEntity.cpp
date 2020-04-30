#include "stdafx.h"
#include "DomeEntity.h"

using namespace OCC;

CDomeEntity::CDomeEntity(void)
{
	m_type = OCC::CDomeEntity::TypeString();

	Reset(NULL);
}

CDomeEntity::~CDomeEntity(void)
{
}

/******************************************************************************
    @brief		return type string of dome
    @author     humkyung
    @date       2014.08.08
    @class      CConeEntity
    @function   TypeString
    @return     STRING_T
******************************************************************************/
STRING_T CDomeEntity::TypeString()
{
	static const STRING_T __type_str__(_T("dome"));
	
	return __type_str__;
}

int CDomeEntity::Reset(Handle_AIS_InteractiveContext hAISContext)
{
	COCCEntity::Reset(hAISContext);

	m_axis = gp_Dir(0.0 , 0.0 , 1.0);
	m_dWidth = 10.0;
	m_dHeight= 7.0;

	return ERROR_SUCCESS;
}

void CDomeEntity::Rotate( const double& angle )
{

}

void CDomeEntity::Rotate( const gp_Ax1& axis , const double& angle )
{
	COCCEntity::Rotate( axis , angle );
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

/******************************************************************************
    @author     humkyung
    @date       2011-09-01
    @class      CDomeEntity
    @function   CreateShape
    @return     int
    @param      list<Handle(AIS_Shape
    @brief
******************************************************************************/
TopoDS_Shape CDomeEntity::CreateShape(BRep_Builder* pBuilder , TopoDS_Compound* pCompound)
{
	gp_Ax2 ax2 = gp_Ax2(m_ptOrigin , m_axis);
	gp_Dir xdir = ax2.XDirection();
	gp_Dir ydir = ax2.YDirection();

	gp_Elips e = gp_Elips(gp_Ax2(m_ptOrigin , m_axis) , m_dWidth , m_dHeight);
	e.Rotate(gp_Ax1(m_ptOrigin , xdir) , 0.5*PI);
	Handle(Geom_TrimmedCurve) aTrimmedCurve = GC_MakeArcOfEllipse(e , 0 , 0.5*PI , true);

	TopoDS_Edge aEdge = BRepBuilderAPI_MakeEdge(aTrimmedCurve);
	TopoDS_Wire aWire = BRepBuilderAPI_MakeWire(aEdge);

	gp_Ax1 axe = gp_Ax1(m_ptOrigin , gp_Dir(m_axis));
	TopoDS_Shape aShape = BRepPrimAPI_MakeRevol(aWire , axe , 360.*PI180);
	return aShape;

	/*if((NULL != pBuilder) && (NULL != pCompound))
	{
		pBuilder->Add( *pCompound , aShape );
	}
	else
	{
		if(m_oAISShapeList.empty())
		{
			Handle(AIS_Shape) aAISShape= new AIS_Shape(aShape);
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

void CDomeEntity::Translate(gp_Vec& V)
{
	m_ptOrigin.Translate(V);	
	if(!m_oAISShapeList.empty())
	{
		CreateShape();
	}
}

void CDomeEntity::Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext)
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

	if(!CreateShape().IsNull()) Redisplay(hContext);
}