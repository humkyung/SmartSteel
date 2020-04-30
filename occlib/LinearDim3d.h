#pragma once

#ifndef _Standard_Macro_HeaderFile
#include <Standard_Macro.hxx>
#endif

#include <AIS_InteractiveObject.hxx>
#include <MMgt_TShared.hxx>
#include <Standard_DefineHandle.hxx>
#include <gp_Pnt.hxx>

// Handle definition
//
DEFINE_STANDARD_HANDLE(CLinearDim3d , AIS_InteractiveObject)

class CLinearDim3d: public AIS_InteractiveObject
{
public:
	Standard_EXPORT CLinearDim3d();
	Standard_EXPORT CLinearDim3d(const Standard_Real& x, const Standard_Real& y, const Standard_Real& z);
	Standard_EXPORT CLinearDim3d(const gp_Pnt& pos );
	Standard_EXPORT CLinearDim3d(const gp_Pnt& origin , const gp_Dir& axis , const gp_Pnt& pos1 , const gp_Pnt& pos2);

	gp_Pnt GetOrigin() const { return m_origin; } 
	void SetOrigin(const gp_Pnt& origin);
	virtual Standard_Real Distance(const gp_Pnt& point) const;

	// some methods like DynamicType() or IsKind()
	DEFINE_STANDARD_RTTI(CLinearDim3d)
private:
	Standard_EXPORT void Compute          (const Handle(PrsMgr_PresentationManager3d)& aPresentationManager,
		const Handle(Prs3d_Presentation)& aPresentation,
		const Standard_Integer aMode);
	void Compute          (const Handle(Prs3d_Projector)& aProjector,
		const Handle(Prs3d_Presentation)& aPresentation);
	/*void Compute          (const Handle(PrsMgr_PresentationManager2d)& aPresentationManager,
		const Handle(Graphic2d_GraphicObject)& aGrObj,
		const Standard_Integer unMode = 0) ;*/
	void ComputeSelection (const Handle(SelectMgr_Selection)& aSelection,
		const Standard_Integer unMode) ;
private:
	gp_Pnt m_origin;
	gp_Dir m_axis;
	gp_Pnt m_position[2];
};