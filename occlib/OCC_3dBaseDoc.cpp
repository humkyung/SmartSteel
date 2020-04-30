// OCC_3dBaseDoc.cpp: implementation of the OCC_3dBaseDoc class.
//
//////////////////////////////////////////////////////////////////////

#include "Stdafx.h"

#include "OCC_3dBaseDoc.h"

#include "OCC_3dView.h"
#include "OCC_3dApp.h"
///#include <res\OCC_Resource.h>
#include "ImportExport/ImportExport.h"
///#include "AISDialogs.h"
#include <AIS_ListOfInteractive.hxx>
#include <AIS_ListIteratorOfListOfInteractive.hxx>

BEGIN_MESSAGE_MAP(OCC_3dBaseDoc, OCC_BaseDoc)
	//{{AFX_MSG_MAP(OCC_3dBaseDoc)
	///ON_COMMAND(ID_FILE_IMPORT_BREP, OnFileImportBrep)
	///ON_COMMAND(ID_FILE_EXPORT_BREP, OnFileExportBrep)
	///ON_COMMAND(ID_OBJECT_ERASE, OnObjectErase)
	///ON_UPDATE_COMMAND_UI(ID_OBJECT_ERASE, OnUpdateObjectErase)
	///ON_COMMAND(ID_OBJECT_COLOR, OnObjectColor)
	///ON_UPDATE_COMMAND_UI(ID_OBJECT_COLOR, OnUpdateObjectColor)
	///ON_COMMAND(ID_OBJECT_SHADING, OnObjectShading)
	//ON_UPDATE_COMMAND_UI(ID_OBJECT_SHADING, OnUpdateObjectShading)
	///ON_COMMAND(ID_OBJECT_WIREFRAME, OnObjectWireframe)
	///ON_UPDATE_COMMAND_UI(ID_OBJECT_WIREFRAME, OnUpdateObjectWireframe)
	///ON_COMMAND(ID_OBJECT_TRANSPARENCY, OnObjectTransparency)
	///ON_UPDATE_COMMAND_UI(ID_OBJECT_TRANSPARENCY, OnUpdateObjectTransparency)
	//ON_COMMAND(ID_OBJECT_MATERIAL, OnObjectMaterial)
	//ON_UPDATE_COMMAND_UI(ID_OBJECT_MATERIAL, OnUpdateObjectMaterial)
	//ON_COMMAND(ID_OBJECT_DISPLAYALL, OnObjectDisplayall)
	//ON_UPDATE_COMMAND_UI(ID_OBJECT_DISPLAYALL, OnUpdateObjectDisplayall)
	//ON_COMMAND(ID_OBJECT_REMOVE, OnObjectRemove)
	//ON_UPDATE_COMMAND_UI(ID_OBJECT_REMOVE, OnUpdateObjectRemove)
	////}}AFX_MSG_MAP
	//ON_COMMAND_EX_RANGE(ID_OBJECT_MATERIAL_BRASS,ID_OBJECT_MATERIAL_DEFAULT, OnObjectMaterialRange)
	//ON_UPDATE_COMMAND_UI_RANGE(ID_OBJECT_MATERIAL_BRASS,ID_OBJECT_MATERIAL_DEFAULT, OnUpdateObjectMaterialRange)

END_MESSAGE_MAP()


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OCC_3dBaseDoc::OCC_3dBaseDoc()
{
	AfxInitRichEdit();

	Handle(Graphic3d_GraphicDriver) theGraphicDevice = 
		((OCC_3dApp*)AfxGetApp())->GetGraphicDevice();

	m_hViewer = new V3d_Viewer(theGraphicDevice,(short *) "Visu3D");
	m_hViewer->SetDefaultLights();
	m_hViewer->SetLightOn();

	// Create an interactive context based on the m_hViewer
	m_hAISContext =new AIS_InteractiveContext(m_hViewer);
}

OCC_3dBaseDoc::~OCC_3dBaseDoc()
{

}


void OCC_3dBaseDoc::DragEvent(const Standard_Integer  x        ,
			      const Standard_Integer  y        ,
			      const Standard_Integer  TheState ,
			      const Handle(V3d_View)& aView    )
{

	// TheState == -1  button down
	// TheState ==  0  move
	// TheState ==  1  button up

	static Standard_Integer theButtonDownX=0;
	static Standard_Integer theButtonDownY=0;

	if (TheState == -1)
	{
		theButtonDownX=x;
		theButtonDownY=y;
	}

	if (TheState == 0)
	{
		m_hAISContext->Select(theButtonDownX,theButtonDownY,x,y,aView);  
	}
}

//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void OCC_3dBaseDoc::InputEvent(const Standard_Integer  x     ,
			       const Standard_Integer  y     ,
			       const Handle(V3d_View)& aView ) 
{
	m_hAISContext->Select(); 
}

//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void OCC_3dBaseDoc::MoveEvent(const Standard_Integer  x       ,
			      const Standard_Integer  y       ,
			      const Handle(V3d_View)& aView   ) 
{
	try
	{
		m_hAISContext->MoveTo(x,y,aView);
	}
	catch(...)
	{
	}
}

//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void OCC_3dBaseDoc::ShiftMoveEvent(const Standard_Integer  x       ,
				   const Standard_Integer  y       ,
				   const Handle(V3d_View)& aView   ) 
{
	m_hAISContext->MoveTo(x,y,aView);
}

//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void OCC_3dBaseDoc::ShiftDragEvent(const Standard_Integer  x        ,
				   const Standard_Integer  y        ,
				   const Standard_Integer  TheState ,
				   const Handle(V3d_View)& aView    ) 
{
	static Standard_Integer theButtonDownX=0;
	static Standard_Integer theButtonDownY=0;

	if (TheState == -1)
	{
		theButtonDownX=x;
		theButtonDownY=y;
	}

	if (TheState == 0)
		m_hAISContext->ShiftSelect(theButtonDownX,theButtonDownY,x,y,aView);  
}


//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void OCC_3dBaseDoc::ShiftInputEvent(const Standard_Integer  x       ,
				    const Standard_Integer  y       ,
				    const Handle(V3d_View)& aView   ) 
{
	m_hAISContext->ShiftSelect(); 
}

//-----------------------------------------------------------------------------------------
//
//-----------------------------------------------------------------------------------------
void  OCC_3dBaseDoc::Popup(const Standard_Integer  x,
			   const Standard_Integer  y ,
			   const Handle(V3d_View)& aView   ) 
{
	//Standard_Integer PopupMenuNumber=0;
	//m_hAISContext->InitCurrent();
	//if (m_hAISContext->MoreCurrent())
	//	PopupMenuNumber=1;

	//CMenu menu;
	//VERIFY(menu.LoadMenu(IDR_Popup3D));
	//CMenu* pPopup = menu.GetSubMenu(PopupMenuNumber);

	//ASSERT(pPopup != NULL);
	//if (PopupMenuNumber == 1) // more than 1 object.
	//{
	//	bool OneOrMoreInShading = false;
	//	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
	//		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) OneOrMoreInShading=true;
	//	if(!OneOrMoreInShading)
	//		pPopup->EnableMenuItem(5, MF_BYPOSITION | MF_DISABLED | MF_GRAYED);
	//}

	//POINT winCoord = { x , y };
	//Handle(WNT_Window) aWNTWindow=
	//	Handle(WNT_Window)::DownCast(aView->Window());
	//ClientToScreen ( (HWND)(aWNTWindow->HWindow()),&winCoord);
	//pPopup->TrackPopupMenu(TPM_LEFTALIGN | TPM_RIGHTBUTTON , winCoord.x, winCoord.y , 
	//	AfxGetMainWnd());
}

void OCC_3dBaseDoc::Fit()
{
	/*
	CMDIFrameWnd *pFrame =  (CMDIFrameWnd*)AfxGetApp()->m_pMainWnd;
	CMDIChildWnd *pChild =  (CMDIChildWnd *) pFrame->GetActiveFrame();
	OCC_3dView *pView = (OCC_3dView *) pChild->GetActiveView();
	pView->FitAll();
	*/
}

int OCC_3dBaseDoc::OnFileImportBrep_WithInitDir(LPCTSTR InitialDir) 
{   
	
	if(CImportExport::ReadBREP(m_hAISContext, InitialDir) == 1)
		return 1;
	Fit();
	
	return 0;
}

void OCC_3dBaseDoc::OnFileImportBrep() 
{   
	/*
	if(CImportExport::ReadBREP(m_hAISContext) == 1)
		return;
	Fit();
	*/
}

void OCC_3dBaseDoc::OnFileExportBrep() 
{   
	///CImportExport::SaveBREP(m_hAISContext);
}

void OCC_3dBaseDoc::OnObjectColor() 
{
	Handle_AIS_InteractiveObject Current ;
	COLORREF MSColor ;
	Quantity_Color CSFColor ;

	m_hAISContext->InitCurrent();
	Current = m_hAISContext->Current();
	if ( Current->HasColor () ) {
		CSFColor = m_hAISContext->Color(m_hAISContext->Current());
		MSColor = RGB (CSFColor.Red()*255.,CSFColor.Green()*255.,CSFColor.Blue()*255.);
	}
	else {
		MSColor = RGB (255,255,255) ;
	}

	CColorDialog dlgColor(MSColor);
	if (dlgColor.DoModal() == IDOK)
	{
		MSColor = dlgColor.GetColor();
		CSFColor = Quantity_Color (GetRValue(MSColor)/255.,GetGValue(MSColor)/255.,
			GetBValue(MSColor)/255.,Quantity_TOC_RGB); 
		for (;m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
			m_hAISContext->SetColor (m_hAISContext->Current(),CSFColor.Name());
	}
}

void OCC_3dBaseDoc::OnUpdateObjectColor(CCmdUI* pCmdUI) 
{
	/*
	bool OneOrMoreIsDisplayed = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
	if (m_hAISContext->IsDisplayed(m_hAISContext->Current())) OneOrMoreIsDisplayed=true;
	pCmdUI->Enable (OneOrMoreIsDisplayed);	
	*/
	bool OneOrMoreIsShadingOrWireframe = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),0)
			||m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) 
			OneOrMoreIsShadingOrWireframe=true;
	pCmdUI->Enable (OneOrMoreIsShadingOrWireframe);	
}

void OCC_3dBaseDoc::OnObjectErase() 
{
	for(m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent();m_hAISContext->NextCurrent())
		m_hAISContext->Erase(m_hAISContext->Current()/*,Standard_True*/,Standard_False);
		m_hAISContext->Erase(m_hAISContext->Current(),Standard_False);
		/*m_hAISContext->Erase(m_hAISContext->Current(),Standard_True,Standard_False);*/
	m_hAISContext->ClearCurrents();
}

void OCC_3dBaseDoc::OnUpdateObjectErase(CCmdUI* pCmdUI) 
{
	bool OneOrMoreIsDisplayed = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current())) OneOrMoreIsDisplayed=true;
	pCmdUI->Enable (OneOrMoreIsDisplayed);	
}

void OCC_3dBaseDoc::OnObjectWireframe() 
{
	for(m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent();m_hAISContext->NextCurrent())
		m_hAISContext->SetDisplayMode(m_hAISContext->Current(),0);
}
void OCC_3dBaseDoc::OnUpdateObjectWireframe(CCmdUI* pCmdUI) 
{
	bool OneOrMoreInShading = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) OneOrMoreInShading=true;
	pCmdUI->Enable (OneOrMoreInShading);	
}

void OCC_3dBaseDoc::OnObjectShading() 
{
	for(m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent();m_hAISContext->NextCurrent())
		m_hAISContext->SetDisplayMode(m_hAISContext->Current(),1);
}

void OCC_3dBaseDoc::OnUpdateObjectShading(CCmdUI* pCmdUI) 
{
	bool OneOrMoreInWireframe = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),0)) OneOrMoreInWireframe=true;
	pCmdUI->Enable (OneOrMoreInWireframe);	
}

void OCC_3dBaseDoc::OnObjectMaterial() 
{
	/*
	CDialogMaterial DialBox(m_hAISContext);
	DialBox.DoModal();
	CMDIFrameWnd *pFrame =  (CMDIFrameWnd*)AfxGetApp()->m_pMainWnd;
	CMDIChildWnd *pChild =  (CMDIChildWnd *) pFrame->GetActiveFrame();
	OCC_3dView *pView = (OCC_3dView *) pChild->GetActiveView();
	pView->Redraw();
	*/
}

void OCC_3dBaseDoc::OnUpdateObjectMaterial(CCmdUI* pCmdUI) 
{
	bool OneOrMoreInShading = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) OneOrMoreInShading=true;
	pCmdUI->Enable (OneOrMoreInShading);	
}

BOOL OCC_3dBaseDoc::OnObjectMaterialRange(UINT nID) 
{
	// the range ID_OBJECT_MATERIAL_BRASS to ID_OBJECT_MATERIAL_SILVER is
	// continue with the same values as enumeration Type Of Material
	/*Standard_Real aTransparency;

	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ()){
		aTransparency = m_hAISContext->Current()->Transparency();
		m_hAISContext->SetMaterial (m_hAISContext->Current(),(Graphic3d_NameOfMaterial)(nID-ID_OBJECT_MATERIAL_BRASS));
		m_hAISContext->SetTransparency (m_hAISContext->Current(),aTransparency);
	}*/
	return true;

}

void OCC_3dBaseDoc::OnUpdateObjectMaterialRange(CCmdUI* pCmdUI) 
{
	//bool OneOrMoreInShading = false;
	//for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
	//	if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) OneOrMoreInShading=true;
	//pCmdUI->Enable (OneOrMoreInShading);
	//for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
	//	if (m_hAISContext->Current()->Material() - (pCmdUI->m_nID - ID_OBJECT_MATERIAL_BRASS) == 0) 
	//		pCmdUI->SetCheck(1);	
}


void OCC_3dBaseDoc::OnObjectTransparency()
{
	/*
	CDialogTransparency DialBox(m_hAISContext);
	DialBox.DoModal();
	CMDIFrameWnd *pFrame =  (CMDIFrameWnd*)AfxGetApp()->m_pMainWnd;
	CMDIChildWnd *pChild =  (CMDIChildWnd *) pFrame->GetActiveFrame();
	OCC_3dView *pView = (OCC_3dView *) pChild->GetActiveView();
	pView->Redraw();
	*/
}

void OCC_3dBaseDoc::OnUpdateObjectTransparency(CCmdUI* pCmdUI) 
{
	bool OneOrMoreInShading = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current(),1)) OneOrMoreInShading=true;
	pCmdUI->Enable (OneOrMoreInShading);	
}


void OCC_3dBaseDoc::OnObjectDisplayall() 
{
	m_hAISContext->DisplayAll(Standard_True);
}

void OCC_3dBaseDoc::OnUpdateObjectDisplayall(CCmdUI* pCmdUI) 
{

	AIS_ListOfInteractive aList;
	m_hAISContext->ObjectsInside(aList,AIS_KOI_Shape);
	AIS_ListIteratorOfListOfInteractive aLI;
	Standard_Boolean IS_ANY_OBJECT_ERASED=FALSE;
	for (aLI.Initialize(aList);aLI.More();aLI.Next()){
		if(!m_hAISContext->IsDisplayed(aLI.Value()))
			IS_ANY_OBJECT_ERASED=TRUE;
	}
	pCmdUI->Enable (IS_ANY_OBJECT_ERASED);

}

void OCC_3dBaseDoc::OnObjectRemove() 
{
	for(m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent();m_hAISContext->InitCurrent())
		m_hAISContext->Remove(m_hAISContext->Current(),Standard_True);
}

void OCC_3dBaseDoc::OnUpdateObjectRemove(CCmdUI* pCmdUI) 
{
	bool OneOrMoreIsDisplayed = false;
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		if (m_hAISContext->IsDisplayed(m_hAISContext->Current())) OneOrMoreIsDisplayed=true;
	pCmdUI->Enable (OneOrMoreIsDisplayed);	
}

void OCC_3dBaseDoc::SetMaterial(Graphic3d_NameOfMaterial Material) 
{
	for (m_hAISContext->InitCurrent();m_hAISContext->MoreCurrent ();m_hAISContext->NextCurrent ())
		m_hAISContext->SetMaterial (m_hAISContext->Current(),
		(Graphic3d_NameOfMaterial)(Material));
}
