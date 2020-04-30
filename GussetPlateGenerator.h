#pragma once

#include <IsPoint3d.h>

#include "GussetPlate.h"
#include "SteelConnPoint.h"
#include "SmartSteelDoc.h"

#include <list>
using namespace std;

class CGussetPlateGenerator
{
	CGussetPlateGenerator(void);
	CGussetPlateGenerator(const CGussetPlateGenerator&){}
	CGussetPlateGenerator& operator =(const CGussetPlateGenerator&){return (*this);}
public:
	static CGussetPlateGenerator& GetInstance(void);
	~CGussetPlateGenerator(void);

	int Generate(list<CGussetPlate*>& , CSteelConnPoint* , CSmartSteelDoc*);

	/**
	@brief	check given two direction has same direction
	*/
	static bool IsSameDir(const CIsVect3d lhs , const CIsVect3d& rhs);

	///check given dir is parallel to width dir of linear member
	static bool IsWebDir(CSDNFLinearMember* pMember , const CIsVect3d& dir);
	
	///check given dir is parallel to height dir of linear member
	static bool IsFlangeDir(CSDNFLinearMember* pMember , const CIsVect3d& dir);
	
	static bool IsEqual(const CIsPoint3d& lhs , const CIsPoint3d& rhs , const double& dToler=0.01);	/// 2013.06.26 added by humkyung

	/**
	@brief	return the offset from linear member depend on connection dir.
	*/
	static double GetOffsetFromLinearMember(CSDNFLinearMember* , const CSDNFLinearMember::ElmType& , const CIsPoint3d& , const CIsVect3d& , CSmartSteelDoc*);
private:		
	bool CheckIfColumnBeamVBraceIsPossible(CSteelConnPoint*) const;	/// 2014.05.21 added by humkyung

	/**
	@brief	generate gusset plate for column to ver. brace
	*/
	int Generate4ColumnToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc);

	/**
	@brief	generate gusset plate for beam to hor. brace
	*/
	int Generate4BeamToHBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*>& oMemList , CSmartSteelDoc* pDoc);

	/**
	@brief	generate gusset plate for beam to ver. brace
	*/
	int Generate4BeamToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc);

	/// generate gusset plate for vbrace to vbrace
	int Generate4VBraceToVBrace(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*> oMemList , CSmartSteelDoc* pDoc);
};
