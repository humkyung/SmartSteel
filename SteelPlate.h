#pragma once

#include <occlib.h>
#include <ComplexShapeEntity.h>
#include "SteelConnPoint.h"

#include <vector>
using namespace std;

class Handle_AIS_InteractiveContext;
class Handle_V3d_View;
class Handle(AIS_InteractiveObject);

class CSteelPlate
{
public:
	typedef enum Status
	{
		ALIVE   = 0x01,
		DELETED = 0x02
	};
	
	CSteelPlate(CSteelConnPoint*);
	virtual ~CSteelPlate(void);

	STRING_T GetTypeString() const;
	long id() const;
	long& id();

	CSteelConnPoint* GetConnPnt();

	Bnd_Box BoundBox() const;

	/// draw dimension of plate
	int DrawDimension(Handle(AIS_InteractiveContext));

	/// select
	int Select(Handle(AIS_InteractiveContext) hContext);

	/// return shape entity list
	vector<OCC::CComplexShapeEntity*>* GetShapeEntList();

	const Status status() const;
	Status& status();

	/// return max edge length
	double GetMaximumEdgeLength() const;
	virtual int Display(Handle(AIS_InteractiveContext) hContext) = 0;
	virtual int Show(Handle(AIS_InteractiveContext) hContext , const bool& bShow) = 0;
	virtual int Write(OFSTREAM_T& ofile , const double dUnitScale) = 0;
protected:
	STRING_T m_sType;
	long m_id;
	CSteelConnPoint* m_pConnPnt;

	vector<CIsPoint3d> m_oSectionShapePntList;
	CIsVect3d m_norm;
	double m_dThickness;

	vector<OCC::CComplexShapeEntity*> m_oShapeEntList;
	Status m_status;
};
