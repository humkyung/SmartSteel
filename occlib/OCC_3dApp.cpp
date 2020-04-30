// OCC_3dApp.cpp: implementation of the OCC_3dApp class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <OSD.hxx>
#include <Standard_Version.hxx>
#include <Graphic3d.hxx>
#include "OCC_3dApp.h"
#include <Aspect_DisplayConnection.hxx>
#include <Graphic3d_GraphicDriver.hxx>

Handle_Graphic3d_GraphicDriver theGraphicDevice;
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OCC_3dApp::OCC_3dApp()
{
	OSD::SetSignal (Standard_True);
	SampleName = "";
	SetSamplePath (NULL);
	try
	{
		Handle(Aspect_DisplayConnection) aDisplayConnection;
		theGraphicDevice = Graphic3d::InitGraphicDriver(aDisplayConnection);
		//m_pGraphicDevice = ::new Graphic3d_GraphicDriver;
	}
	catch(Standard_Failure a)
	{
		AfxMessageBox( _T("Fatal Error During Graphic Initialisation") );
		ExitProcess(1);
	};

	// Set the local system units
	/*try
	{
		UnitsAPI::SetLocalSystem(UnitsAPI_DEFAULT);
	}
	catch (Standard_Failure)
	{
		AfxMessageBox("Fatal Error in units initialisation");
	}
	*/
}

OCC_3dApp::~OCC_3dApp()
{

}
