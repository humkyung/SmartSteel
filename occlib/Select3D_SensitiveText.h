#pragma once

#ifndef _Standard_HeaderFile
#include <Standard.hxx>
#endif
#ifndef _Handle_Select3D_SensitivePoint_HeaderFile
#include <Handle_Select3D_SensitivePoint.hxx>
#endif

#ifndef _Select3D_Pnt_HeaderFile
#include <Select3D_Pnt.hxx>
#endif
#ifndef _Select3D_Pnt2d_HeaderFile
#include <Select3D_Pnt2d.hxx>
#endif
#ifndef _Select3D_SensitiveEntity_HeaderFile
#include <Select3D_SensitiveEntity.hxx>
#endif
#ifndef _Handle_SelectBasics_EntityOwner_HeaderFile
#include <Handle_SelectBasics_EntityOwner.hxx>
#endif
#ifndef _Handle_Select3D_SensitiveEntity_HeaderFile
#include <Handle_Select3D_SensitiveEntity.hxx>
#endif
#ifndef _Standard_Boolean_HeaderFile
#include <Standard_Boolean.hxx>
#endif
#ifndef _Standard_Real_HeaderFile
#include <Standard_Real.hxx>
#endif
#ifndef _Standard_OStream_HeaderFile
#include <Standard_OStream.hxx>
#endif

class SelectBasics_EntityOwner;
class gp_Pnt;
class Select3D_Projector;
class SelectBasics_ListOfBox2d;
class Select3D_SensitiveEntity;
class TopLoc_Location;
class TColgp_Array1OfPnt2d;
class Bnd_Box2d;
class gp_Lin;

class Select3D_SensitiveText : public Select3D_SensitiveEntity
{
public:
	Standard_EXPORT Select3D_SensitiveText(const Handle(SelectBasics_EntityOwner)& OwnerId);
	~Select3D_SensitiveText(void);
};
