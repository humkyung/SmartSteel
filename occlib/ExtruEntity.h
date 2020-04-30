#pragma once
#include "occentity.h"

#include <vector>
using namespace std;

namespace OCC
{
class AFX_EXT_CLASS CExtruEntity : public COCCEntity
{
public:
	CExtruEntity(void);
	~CExtruEntity(void);

	void Rotate( const double& angle );
	void Rotate( const gp_Ax1& axis , const double& angle );
	int Reset(Handle_AIS_InteractiveContext hAISContext);
	TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL);
	void Translate(gp_Vec& V);
	void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);

	static STRING_T TypeString();
public:
	TopoDS_Wire m_hStartWire , m_hEndWire;
	vector<gp_Pnt> m_oPointList;
	gp_Dir m_axis;
	double m_dThickness;
};
};