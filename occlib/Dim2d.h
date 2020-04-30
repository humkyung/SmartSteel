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

#include <MMgt_TShared.hxx>
#include <Standard_DefineHandle.hxx>
#include <gp_Pnt2d.hxx>

// Handle definition
//
DEFINE_STANDARD_HANDLE(CDim2d , MMgt_TShared)

class CDim2d: public MMgt_TShared
{
public:
	Standard_EXPORT CDim2d();
	Standard_EXPORT CDim2d(const Standard_Real& x, const Standard_Real& y);
	Standard_EXPORT CDim2d(const gp_Pnt2d& pos);

	gp_Pnt2d GetPoint() const { return position; } 
	void SetPoint(const gp_Pnt2d& pos); 
	void SetCoord(const Standard_Real& x, const Standard_Real& y);
	virtual Standard_Real Distance(const gp_Pnt2d& point) const;

	// some methods like DynamicType() or IsKind()
	DEFINE_STANDARD_RTTI(CDim2d)

private:
	gp_Pnt2d position;
};