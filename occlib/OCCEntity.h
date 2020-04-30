#pragma once

#include "StdAfx.h"
//#include <AIS_ColoredShape.hxx>
#include "OCCAttribute.h"
#include "Standard_StdAllocator.hxx"

#include <list>
#include <vector>
using namespace std;

typedef Standard_StdAllocator<Handle(AIS_Shape)> allocator_Handle_AIS_Shape;
namespace OCC
{
#define	PI180	(3.14/180.0)

class AFX_EXT_CLASS COCCEntity
{
public:
	COCCEntity(void);
	virtual ~COCCEntity(void);

	/// add AIS_Shape in entity to OpenCASCADE selection set
	void Select(Handle(AIS_InteractiveContext) hContext);

	void SetSolid(const bool& bSolid);	/// 2011.12.08 added by humkyung

	static TopoDS_Wire CreateWire(const vector<gp_Pnt>& oPointList);	/// 2011.12.07 added by humkyung

	virtual Bnd_Box BoundingBox();	/// 2011.12.04 added by humkyung

	virtual int Show(Handle_AIS_InteractiveContext hContext , const bool& bShow);	/// 2011.11.22 added by humkyung
	int AddAttribute(COCCAttribute* pAttr);			/// 2011.11.22 added by humkyung
	COCCAttribute* GetAttributeAt( const int& at );	/// 2011.11.22 added by humkyung
	int GetAttributeCount() const;					/// 2011.11.22 added by humkyung

	virtual int SetTransparency( const double dTransparency );	/// 2011.11.21 added by humkyung

	gp_Pnt GetOrigin() const;		/// 2011.11.18 added by humkyung
	int SetOrigin(gp_Pnt& origin);	/// 2011.11.18 added by humkyung

	Handle_AIS_InteractiveContext GetContext();	/// 2011.10.20 added by humkyung

	virtual COCCEntity* Clone();
	virtual int Reset(Handle_AIS_InteractiveContext hAISContext) = 0;
	virtual TopoDS_Shape CreateShape(BRep_Builder* pBuilder = NULL , TopoDS_Compound* pCompound = NULL) = 0;
	virtual void Display(Handle_AIS_InteractiveContext hContext , const Standard_Integer& aMode);
	void Redisplay(Handle_AIS_InteractiveContext hContext);
	virtual void Rotate( const double& angle );
	virtual void Rotate( const gp_Ax1& axis , const double& angle );
	virtual void Translate(Handle_AIS_InteractiveContext hContext , gp_Vec& V);
	virtual void Update(vector<CString>& oObject , Handle_AIS_InteractiveContext hContext);
	Quantity_NameOfColor GetColorFrom(const CString& sColor);

	int SetColor(const CString& sColor);
	CString guid() const;
	STRING_T type() const;

	static double EPSILON;
protected:
	STRING_T m_type;

	CString m_guid;
	gp_Pnt m_ptOrigin;
	CString m_sColor;
	bool m_bSolid;

	double m_dTransparency;
	
	vector<COCCAttribute*> m_oAttributeList;
	vector<Handle(AIS_Shape) , allocator_Handle_AIS_Shape> m_oAISShapeList;
};
};