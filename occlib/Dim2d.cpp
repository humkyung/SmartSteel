#include "Dim2d.h"

// Implementation of Handle and type mgt
//
IMPLEMENT_STANDARD_HANDLE(CDim2d , MMgt_TShared)
IMPLEMENT_STANDARD_RTTI(CDim2d)

IMPLEMENT_STANDARD_TYPE(CDim2d)
	IMPLEMENT_STANDARD_SUPERTYPE(MMgt_TShared)
	IMPLEMENT_STANDARD_SUPERTYPE(Standard_Transient)
	IMPLEMENT_STANDARD_SUPERTYPE_ARRAY()
	IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(MMgt_TShared)
	IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_ENTRY(Standard_Transient)
	IMPLEMENT_STANDARD_SUPERTYPE_ARRAY_END()
IMPLEMENT_STANDARD_TYPE_END(CDim2d)

// Constructors implementation
//
CDim2d::CDim2d() : position(0,0)
{}

CDim2d::CDim2d(const Standard_Real& x, const Standard_Real& y) : position(x,y)
{}

CDim2d::CDim2d(const gp_Pnt2d& pos) : position(pos)
{}

void CDim2d::SetPoint(const gp_Pnt2d& pos)
{
	position = pos;
}

void CDim2d::SetCoord(const Standard_Real& x, const Standard_Real& y)
{
	position = gp_Pnt2d(x,y);
}

Standard_Real CDim2d::Distance(const gp_Pnt2d& point) const
{
	return position.Distance(point);
}
