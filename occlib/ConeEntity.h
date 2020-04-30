#pragma once
#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CConeEntity : public COCCEntity
{
public:
	CConeEntity(void);
	~CConeEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Translate(gp_Vec& V);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);

	static STRING_T TypeString();
public:
	gp_Pnt m_ptBottom , m_ptTop;
	double m_dBottomRadius , m_dTopRadius;
	gp_Dir m_axis;
private:
};
};