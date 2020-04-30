#include "StdAfx.h"
#include <tchar.h>
#include <gl/gl.h>

#include <Standard_Macro.hxx>
//#include <Graphic3d_NameOfFont.hxx>
#include <gp_Circ.hxx>
#include <GC_MakeArcOfCircle.hxx>
#include <AIS_Line.hxx>
#include <GC_MakeLine.hxx>
#include <Prs3d_TextAspect.hxx>
#include <Prs3d_Arrow.hxx>

#include <StdPrs_Curve.hxx>
#include <Geom_Line.hxx>

#include "OCCEntity.h"
#include "AngularDim3d.h"

// Implementation of Handle and type mgt
//
IMPLEMENT_STANDARD_HANDLE(CAngularDim3d , AIS_InteractiveObject)
IMPLEMENT_STANDARD_RTTIEXT(CAngularDim3d , AIS_InteractiveObject)

//IMPLEMENT_STANDARD_TYPE(CAngularDim3d)
//IMPLEMENT_STANDARD_SUPERTYPE(AIS_InteractiveObject)
//IMPLEMENT_STANDARD_SUPERTYPE(SelectMgr_SelectableObject)
//IMPLEMENT_STANDARD_SUPERTYPE(PrsMgr_PresentableObject)
//IMPLEMENT_STANDARD_SUPERTYPE(MMgt_TShared)
//IMPLEMENT_STANDARD_SUPERTYPE(Standard_Transient)
//
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY()
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(AIS_InteractiveObject)
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(SelectMgr_SelectableObject)
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(PrsMgr_PresentableObject)
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(MMgt_TShared)
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(Standard_Transient)
//IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_END()
//IMPLEMENT_STANDARD_TYPE_END(CAngularDim3d)

// Constructors implementation
//
CAngularDim3d::CAngularDim3d() : m_origin(0,0,0)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CAngularDim3d::CAngularDim3d(const Standard_Real& x, const Standard_Real& y , const Standard_Real& z) : m_origin(x,y,z)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CAngularDim3d::CAngularDim3d(const gp_Pnt& pos) : m_origin(pos)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CAngularDim3d::CAngularDim3d(const gp_Pnt& origin , const gp_Pnt& pos1 , const gp_Pnt& pos2)
{
	m_origin = origin;
	m_position[0] = pos1;
	m_position[1] = pos2;
}

void CAngularDim3d::SetOrigin(const gp_Pnt& pos)
{
	m_origin = pos;
}

//void CAngularDim3d::SetCoord(const Standard_Real& x, const Standard_Real& y, const Standard_Real& z)
//{
//	m_origin = gp_Pnt(x,y,z);
//}

Standard_Real CAngularDim3d::Distance(const gp_Pnt& point) const
{
	return m_origin.Distance(point);
}

void CAngularDim3d::Compute(const Handle(PrsMgr_PresentationManager3d)& aPresentationManager,
								   const Handle(Prs3d_Presentation)& aPresentation,
								   const Standard_Integer aMode)
{
	gp_Pnt origin(m_origin);
	origin.SetZ( m_position[0].Z());
	if(!origin.IsEqual(m_position[0] , OCC::COCCEntity::EPSILON))
	{
		gp_XYZ xyz = (m_position[0].XYZ() - origin.XYZ());
		xyz.Normalize();
		gp_Dir dir(xyz);

		Handle(Geom_Line) aComponent = new Geom_Line(origin , dir);
		const Standard_Real dist = origin.Distance(m_position[0]);
		GeomAdaptor_Curve curv(aComponent , 0. , dist);

		bool bHasArcDim = false;
		gp_Dir aVec = (m_position[0].Y() < origin.Y()) ? gp_Dir(0,0,-1) : gp_Dir(0,0,1);
		gp_Circ c = gp_Circ(gp_Ax2(origin , aVec/*m_position[1].XYZ() - m_position[0].XYZ())*/) , dist);
		gp_Pnt ptStart(origin.X() + dist , origin.Y() , origin.Z());
		if(!ptStart.IsEqual(m_position[0] , OCC::COCCEntity::EPSILON))
		{
			Handle(Geom_TrimmedCurve) aTrimmedCurve = GC_MakeArcOfCircle(c , ptStart , m_position[0] , true);
			GeomAdaptor_Curve curve2(aTrimmedCurve);

			myDrawer->LineAspect()->SetColor(Quantity_NOC_RED);   
			myDrawer->TextAspect()->SetFont("Courier New"/*Graphic3d_NOF_ASCII_MONO*/);
			switch (aMode)
			{   
			case 1 :   
				StdPrs_Curve::Add(aPresentation, curv , myDrawer);   
				break;
				///StdPrs_PoleCurve::Add(aPresentation, anAdaptorCurve2 , aDrawer);   
			case 0 :   
				StdPrs_Curve::Add(aPresentation, curv , myDrawer);
				///StdPrs_Curve::Add(aPresentation, curve2 , myDrawer);
				
				Prs3d_Arrow::Draw(aPresentation , m_position[0]  , dir , 0.20 , 0.1);
				break;   
			}  

			/// calculate dim. text position - 2012.01.19 added by humkyung
			gp_Pnt mid((ptStart.X() + m_position[0].X())*0.5 , (ptStart.Y() + m_position[0].Y())*0.5 , (ptStart.Z() + m_position[0].Z())*0.5);
			gp_XYZ dir(m_position[0].XYZ() - origin.XYZ());
			dir.Normalize();
			dir *= (dist*0.5);
			mid = origin.XYZ() + dir;
			/// up to here

			/// calculate angle - 2012.01.19 added by humkyung
			Standard_Real rad = gp_Dir(ptStart.XYZ() - origin.XYZ()).Angle(gp_Dir(m_position[0].XYZ() - origin.XYZ()));
			Standard_Real deg = RAD2DEG(rad);
			deg = SAFE_ROUND(deg , 0);
			OSTRINGSTREAM_T oss;
			oss << deg << _T("deg");
			/// up to here

			Prs3d_Text::Draw(aPresentation , myDrawer , oss.str().c_str() , mid);
		}
	}
}

/// for HLR
void CAngularDim3d::Compute(const Handle(Prs3d_Projector)& aProjector,
								   const Handle(Prs3d_Presentation)& aPresentation) 
{
}

/// for 2D
//void CAngularDim3d::Compute(const Handle(PrsMgr_PresentationManager2d)& aPresentationManager, 
//								   const Handle(Graphic2d_GraphicObject)& aGrObj, 
//								   const Standard_Integer unMode)
//{
//	/*new AIS_Line(m_origin , m_position[0]);
//	new AIS_Line(m_origin , m_position[1]);*/
//	/*Handle(Graphic2d_Text) text;
//	text = new Graphic2d_Text(aGrObj, MyText, MyX, MyY, MyAngle,MyTypeOfText,MyScale);
//	text->SetFontIndex(MyFontIndex);
//
//	text->SetColorIndex(MyColorIndex);
//
//	text->SetSlant(MySlant);
//	text->SetUnderline(Standard_False);
//	text->SetZoomable(Standard_True);
//	aGrObj->Display();
//	Quantity_Length anXoffset,anYoffset;
//	text->TextSize(MyWidth, MyHeight,anXoffset,anYoffset);*/
//}

/// for selection
void CAngularDim3d::ComputeSelection(const Handle(SelectMgr_Selection)& aSelection, 
											const Standard_Integer unMode)
{
	/*
	Handle(SelectMgr_EntityOwner) eown = new SelectMgr_EntityOwner(this);
	Handle(Select3D_SensitiveText) aSensitiveText3d = new Select3D_SensitiveText(eown,5);
	aSelection->Add(aSensitiveText3d);
	*/
}