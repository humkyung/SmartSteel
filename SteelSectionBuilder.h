#pragma once

#include <IsPoint3d.h>
#include <SDNFLinearMember.h>

#include <vector>
using namespace std;

class CSmartSteelDoc;
class CSteelSectionBuilder
{
	CSteelSectionBuilder(const CSteelSectionBuilder&){}
	CSteelSectionBuilder& operator=(const CSteelSectionBuilder&){ return (*this);}
public:
	typedef struct
	{
		STRING_T Shape;
		double Height , Width;
		double t1 , t2;
	}ShapeParam;

	CSteelSectionBuilder(void);
	~CSteelSectionBuilder(void);

	int Build(CSDNFLinearMember* pMember , CSmartSteelDoc* pDoc);
private:
	/// 2013.06.24 added by humkyung
	int RotateSection(const CIsVect3d& norm , const double& dRotation);
	int AdjustCardinalPoint(CSDNFLinearMember* pMember , const CIsVect3d& , const double& dWidth , const CIsVect3d& , const double& dHeight);
	/// up to here
public:
	vector<CIsPoint3d> m_oSectionPntList;
	CIsVect3d m_norm;
	double m_thickness;
};
