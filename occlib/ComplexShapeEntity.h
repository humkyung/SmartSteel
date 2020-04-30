#pragma once

#include "OCCEntity.h"

namespace OCC
{
class AFX_EXT_CLASS CComplexShapeEntity : public COCCEntity
{
public:
	CComplexShapeEntity(void);
	~CComplexShapeEntity(void);

	static STRING_T TypeString();

	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	int Reset(Handle_AIS_InteractiveContext hAISContext);
public:
	TopoDS_Shape m_hShape;
};
};
