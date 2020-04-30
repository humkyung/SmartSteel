#include "StdAfx.h"
#include <tchar.h>
#include <gl/gl.h>

#include <gp_Lin2d.hxx>
#include <gp_Circ.hxx>
#include <GC_MakeArcOfCircle.hxx>
#include <AIS_Line.hxx>
#include <GC_MakeLine.hxx>
#include <Prs3d_Arrow.hxx>

#include <StdPrs_Curve.hxx>
#include <Geom_Line.hxx>

#include "OCCEntity.h"
#include "LinearDim3d.h"

// Implementation of Handle and type mgt
//
IMPLEMENT_STANDARD_HANDLE(CLinearDim3d , AIS_InteractiveObject)
IMPLEMENT_STANDARD_RTTIEXT(CLinearDim3d , AIS_InteractiveObject)

//IMPLEMENT_STANDARD_TYPE(CLinearDim3d)
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
//IMPLEMENT_STANDARD_TYPE_END(CLinearDim3d)

// Constructors implementation
//
CLinearDim3d::CLinearDim3d() : AIS_InteractiveObject(PrsMgr_TOP_ProjectorDependant) , m_origin(0,0,0)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CLinearDim3d::CLinearDim3d(const Standard_Real& x, const Standard_Real& y , const Standard_Real& z) : m_origin(x,y,z)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CLinearDim3d::CLinearDim3d(const gp_Pnt& pos) : m_origin(pos)
{
	m_position[0].SetCoord(0,0,0);
	m_position[1].SetCoord(0,0,0);
}

CLinearDim3d::CLinearDim3d(const gp_Pnt& origin , const gp_Dir& axis , const gp_Pnt& pos1 , const gp_Pnt& pos2)
{
	m_origin = origin;
	m_axis = axis;
	m_position[0] = pos1;
	m_position[1] = pos2;
}

void CLinearDim3d::SetOrigin(const gp_Pnt& pos)
{
	m_origin = pos;
}

//void CLinearDim3d::SetCoord(const Standard_Real& x, const Standard_Real& y, const Standard_Real& z)
//{
//	m_origin = gp_Pnt(x,y,z);
//}

Standard_Real CLinearDim3d::Distance(const gp_Pnt& point) const
{
	return m_origin.Distance(point);
}

void CLinearDim3d::Compute(const Handle(PrsMgr_PresentationManager3d)& aPresentationManager,
								   const Handle(Prs3d_Presentation)& aPresentation,
								   const Standard_Integer aMode)
{
	/*gp_Lin2d line2d(gp_Pnt2d(m_position[0].X() , m_position[0].Y()) , 
		gp_Dir2d(m_position[1].X() - m_position[0].X() , m_position[1].Y() - m_position[0].Y()));
	const Standard_Real d = line2d.Distance(gp_Pnt2d(m_origin.X() , m_origin.Y()));
	if(d > 0.0001)
	{
		Handle(Geom_Line) aComponent1 = new Geom_Line(m_position[0] , gp_Dir(m_position[0].XYZ() - m_position[1].XYZ()));
		const Standard_Real length = gp_Pnt2d(m_position[0].X() , m_position[0].Y()).Distance(gp_Pnt2d(m_origin.X() , m_origin.Y()));///m_position[0].Distance(m_position[1]);
		const Standard_Real dimLen = sqrt(length*length - d*d);
		GeomAdaptor_Curve curv1(aComponent1 , 0. , dimLen);

		gp_XYZ xyz1 = m_position[1].XYZ() - m_position[0].XYZ();
		xyz1.SetZ(0.0);
		xyz1.Normalize();
		gp_XYZ xyz2 = m_position[1].XYZ() - m_origin.XYZ();
		xyz2.SetZ(0.0);
		xyz2.Normalize();
		
		gp_Dir axis = xyz2.Crossed(xyz1);
		gp_Dir norm = axis.Crossed(xyz1);

		///gp_Dir norm = m_axis.Crossed(gp_Dir(xyz));
		///norm.Reverse();
		Handle(Geom_Line) aComponent2 = new Geom_Line(m_position[0] , norm);
		GeomAdaptor_Curve curv2(aComponent2 , 0. , d);

		myDrawer->LineAspect()->SetColor(Quantity_NOC_RED);   

		switch (aMode)
		{   
			case 1 :   
				StdPrs_Curve::Add(aPresentation, curv1 , myDrawer);   
				break;
				///StdPrs_PoleCurve::Add(aPresentation, anAdaptorCurve2 , aDrawer);   
			case 0 :   
				StdPrs_Curve::Add(aPresentation, curv1 , myDrawer);
				StdPrs_Curve::Add(aPresentation, curv2 , myDrawer);
				break;  
		}  

		/// calculate dim. text position - 2012.01.19 added by humkyung
		gp_XYZ xyz = norm.XYZ();
		xyz.Normalize();
		gp_Pnt mid(m_position[0].X() + xyz.X()*d*0.5 , m_position[0].Y() + xyz.Y()*d*0.5 , m_position[0].Z() + xyz.Z()*d*0.5);
		gp_Pnt end(m_position[0].X() + xyz.X()*d , m_position[0].Y() + xyz.Y()*d , m_position[0].Z() + xyz.Z()*d);
		/// up to here

		/// display arrows
		Prs3d_Arrow::Draw(aPresentation , m_position[0]  , -norm , 0.20 , d / 10.0);
		Prs3d_Arrow::Draw(aPresentation , end , norm , 0.20 , d / 10.0);
		/// up to here

		Standard_Real value = d;///SAFE_ROUND(d , 0);
		OSTRINGSTREAM_T oss;
		oss << value;

		Prs3d_Text::Draw(aPresentation , myDrawer , oss.str().c_str() , mid);
	}*/

	double dx = m_position[1].X() - m_position[0].X();
	double dy = m_position[1].Y() - m_position[0].Y();
	double dz = m_position[1].Z() - m_position[0].Z();
	gp_Pnt mid(m_position[0].X() + dx*0.5 , m_position[0].Y() + dy*0.5 , m_position[0].Z() + dz*0.5);
	Standard_Real value = sqrt(dx*dx + dy*dy + dz*dz);
	OSTRINGSTREAM_T oss;
	oss << value;

	Prs3d_Text::Draw(aPresentation , myDrawer , oss.str().c_str() , mid);
}

/// for HLR
void CLinearDim3d::Compute(const Handle(Prs3d_Projector)& aProjector,
								   const Handle(Prs3d_Presentation)& aPresentation) 
{
}

/// for 2D
//void CLinearDim3d::Compute(const Handle(PrsMgr_PresentationManager2d)& aPresentationManager, 
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
void CLinearDim3d::ComputeSelection(const Handle(SelectMgr_Selection)& aSelection, 
											const Standard_Integer unMode)
{
	/*
	Handle(SelectMgr_EntityOwner) eown = new SelectMgr_EntityOwner(this);
	Handle(Select3D_SensitiveText) aSensitiveText3d = new Select3D_SensitiveText(eown,5);
	aSelection->Add(aSensitiveText3d);
	*/
}