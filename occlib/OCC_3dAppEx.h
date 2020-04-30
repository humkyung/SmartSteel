// OCC_3dApp.h: interface for the OCC_3dApp class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_OCC_3DAPP_H__FC7278BF_390D_11D7_8611_0060B0EE281E__INCLUDED_)
#define AFX_OCC_3DAPP_H__FC7278BF_390D_11D7_8611_0060B0EE281E__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "StdAfx.h"
#include "OCC_BaseAppEx.h"
#include <Standard_Macro.hxx>
#include <Standard_Failure.hxx>
#include <Graphic3d_GraphicDriver.hxx>

class AFX_EXT_CLASS OCC_3dAppEx : public OCC_BaseAppEx
{
public:
	OCC_3dAppEx();
	virtual ~OCC_3dAppEx();

	Handle_Graphic3d_GraphicDriver GetGraphicDevice() const { return theGraphicDevice; } ;

protected :
};

#endif // !defined(AFX_OCC_3DAPP_H__FC7278BF_390D_11D7_8611_0060B0EE281E__INCLUDED_)
