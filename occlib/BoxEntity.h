#pragma once
#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CBoxEntity : public COCCEntity
{
public:
	CBoxEntity(void);
	~CBoxEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);
public:
	gp_Dir m_xAxis , m_yAxis;
	double m_dWidth , m_dHeight , m_dDepth;
};
};