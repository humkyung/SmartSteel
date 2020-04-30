#pragma once
#include "occentity.h"

namespace OCC
{
class AFX_EXT_CLASS CPyramidEntity : public COCCEntity
{
public:
	CPyramidEntity(void);
	~CPyramidEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Translate(gp_Vec& V);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);

	static STRING_T TypeString();
public:
	gp_Pnt m_ptTop , m_ptBottom;
	double m_dWidth1 , m_dHeight1;
	double m_dWidth2 , m_dHeight2;
	gp_Dir m_xAxis , m_yAxis;
};
};