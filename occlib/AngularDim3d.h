//dim2d_entity.h
//
// (C) 2000  C R Johnson   iamcliff@freeengineer.org
//
// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU  General Public
// License as published by the Free Software Foundation; either
// version 2 of the License, or (at your option) any later version.
//                                                                
// This software is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of 
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// General Public License for more details. 
//                                      
// You should have received a copy of the GNU General Public License 
// along with this software (see COPYING); if not, write to the 
// Free Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.


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
DEFINE_STANDARD_HANDLE(CAngularDim3d , AIS_InteractiveObject)

class CAngularDim3d: public AIS_InteractiveObject
{
public:
	Standard_EXPORT CAngularDim3d();
	Standard_EXPORT CAngularDim3d(const Standard_Real& x, const Standard_Real& y, const Standard_Real& z);
	Standard_EXPORT CAngularDim3d(const gp_Pnt& pos );
	Standard_EXPORT CAngularDim3d(const gp_Pnt& origin , const gp_Pnt& pos1 , const gp_Pnt& pos2);

	gp_Pnt GetOrigin() const { return m_origin; } 
	void SetOrigin(const gp_Pnt& origin);
	virtual Standard_Real Distance(const gp_Pnt& point) const;

	// some methods like DynamicType() or IsKind()
	DEFINE_STANDARD_RTTI(CAngularDim3d)
private:
	void Compute          (const Handle(PrsMgr_PresentationManager3d)& aPresentationManager,
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
	gp_Pnt m_position[2];
};