#pragma once

#include <vector>
#include <IsPoint3d.h>

class CHorGussetPlateGenerator
{
public:
	CHorGussetPlateGenerator(void);
	~CHorGussetPlateGenerator(void);

	/*
	@brief	generate horizontal guesset plate
	@param	oGussetPlateList
	@param	pConnPnt
	@param	pDoc
	*/
	int Generate(list<CGussetPlate*>& oGussetPlateList , CSteelConnPoint* pConnPnt , CSmartSteelDoc* pDoc);
	int GetOuterPntListOfPlate(vector<CIsPoint3d>& oShapePntList , CIsVect3d& norm , CSteelConnPoint* pConnPnt , vector<CSDNFLinearMember*>& oMemList , CSmartSteelDoc* pDoc);
private:
	/*
	@brief	find a base linear member
	@param	pConnPnt
	*/
	CSDNFLinearMember* FindBaseMember(CSteelConnPoint* pConnPnt);
};
