#pragma once

#include <Standard.hxx>
#include <Standard_PrimitiveTypes.hxx>

#include <AIS_InteractiveContext.hxx>
#include <AIS_Line.hxx>
#include <AIS_Shape.hxx>
#include <AIS_Drawer.hxx>
#include <AIS_Point.hxx>
#include <AIS_TexturedShape.hxx>
#include <Aspect_Grid.hxx>
#include <Aspect_PolygonOffsetMode.hxx>
#include <Aspect_TypeOfText.hxx>
#include <Aspect_DisplayConnection.hxx>
#include <Aspect_AspectMarker.hxx>

#include <BRep_Tool.hxx>
#include <BRepTools.hxx>
#include <BRepBuilderAPI_NurbsConvert.hxx>
#include <BRepBuilderAPI_MakeEdge.hxx>
#include <BRepBuilderAPI_MakeWire.hxx>
#include <BRepBuilderAPI_MakeFace.hxx>
#include <BRepBndLib.hxx>
#include <BRepAdaptor_HArray1OfCurve.hxx>
#include <BRepAdaptor_Curve2d.hxx>
#include <BRepBuilderAPI_MakeVertex.hxx>
#include <BRepOffsetAPI_ThruSections.hxx>

#include <BRepPrimAPI_MakeCylinder.hxx>
#include <BRepPrimAPI_MakeBox.hxx>
#include <BRepPrimAPI_MakeRevol.hxx>
#include <BRepPrimAPI_MakeSphere.hxx>

#include <GC_MakeArcOfEllipse.hxx>

#include <Geom2d_TrimmedCurve.hxx>
#include <GeomLib.hxx>
#include <Geom_Surface.hxx>
#include <Geom_Curve.hxx>
#include <Geom_Plane.hxx>
#include <Geom_CartesianPoint.hxx>
#include <Geom_TrimmedCurve.hxx>

#include <GCPnts_QuasiUniformDeflection.hxx>
#include <Graphic3d_Group.hxx>
#include <Graphic3d_HorizontalTextAlignment.hxx>
#include <Graphic3d_VerticalTextAlignment.hxx>
#include <Graphic3d_Array1OfVertex.hxx>
#include <Graphic3d_ArrayOfPolylines.hxx>
//#include <Graphic3d.hxx>
#include <Graphic3d_ExportFormat.hxx>
#include <Graphic3d_ArrayOfPolylines.hxx>
#include <Graphic3d_AspectFillArea3d.hxx>
#include <Graphic3d_AspectText3d.hxx>
#include <Graphic3d_AspectLine3d.hxx>
#include <Graphic3d_AspectMarker3d.hxx>
#include <Graphic3d_Texture1Dsegment.hxx>

#include <gp_Pln.hxx>
#include <gp.hxx>
#include <gp_Pnt2d.hxx>
#include <gp_Elips.hxx>

#include <Prs3d_Root.hxx>
#include <Prs3d_Drawer.hxx>
#include <Prs3d_IsoAspect.hxx>
#include <Prs3d_ShadingAspect.hxx>
#include <Prs3d_Presentation.hxx>
#include <PrsMgr_PresentationManager3d.hxx>
#include <Prs3d_TextAspect.hxx>
#include <Prs3d_Text.hxx>

#include <Select3D_ListOfSensitive.hxx>
#include <Select3D_SensitiveBox.hxx>
#include <Select3D_SensitiveCurve.hxx>
#include <Select3D_SensitiveGroup.hxx>
#include <SelectMgr_Selection.hxx>
#include <SelectMgr_SequenceOfOwner.hxx>
#include <SelectMgr_EntityOwner.hxx>
#include <ShapeBuild_Edge.hxx>
#include <StdSelect_ViewerSelector3d.hxx>
#include <StdPrs_ShadedShape.hxx>
#include <StdPrs_HLRPolyShape.hxx>
#include <StdSelect_BRepSelectionTool.hxx>
#include <StdPrs_WFDeflectionShape.hxx>
#include <StdPrs_WFShape.hxx>
#include <StdPrs_ToolRFace.hxx>
#include <StdSelect.hxx>
#include <StdSelect_BRepOwner.hxx>
#include <StdSelect_BRepSelectionTool.hxx>

#include <TCollection_AsciiString.hxx>
#include "TopExp.hxx"
#include <TopExp_Explorer.hxx>
#include <TopoDS.hxx>
#include <TopoDS_Builder.hxx>
#include <TopoDS_Compound.hxx>
#include <TopoDS_ListOfShape.hxx>
#include <TopoDS_ListIteratorOfListOfShape.hxx>
#include <TopoDS_Iterator.hxx>
#include "TopoDS_Edge.hxx"
#include "TopoDS_Vertex.hxx"
#include <TopTools_HSequenceOfShape.hxx>
#include <TopTools_IndexedMapOfShape.hxx>

#include <Visual3d_View.hxx>
#include <V3d_Viewer.hxx>
#include <V3d_View.hxx>
#include <WNT_Window.hxx>

#pragma message ("============== Set libs for OpenCASCADE 6.7.1 ==============") 

#pragma comment (lib , "PTKernel.lib")

#pragma comment (lib , "TKOpenGl.lib") 
#pragma comment (lib , "TKBool.lib") 
#pragma comment (lib , "TKBO.lib") 
#pragma comment (lib , "TKVrml.lib")
#pragma comment (lib , "TKStl.lib")
#pragma comment (lib , "TKBRep.lib") 
#pragma comment (lib , "TKIGES.lib")
#pragma comment (lib , "TKStep.lib")
#pragma comment (lib , "TKShapeSchema.lib")
#pragma comment (lib , "TKCAF.lib") 
#pragma comment (lib , "TKCDF.lib") 
///#pragma comment (lib , "TKDraw.lib") 
#pragma comment (lib , "TKernel.lib") 
#pragma comment (lib , "TKFeat.lib") 
#pragma comment (lib , "TKG2d.lib") 
#pragma comment (lib , "TKG3d.lib") 
#pragma comment (lib , "TKGeomAlgo.lib") 
#pragma comment (lib , "TKGeomBase.lib") 
#pragma comment (lib , "TKMath.lib") 
#pragma comment (lib , "TKOffset.lib") 
#pragma comment (lib , "TKPCAF.lib") 
#pragma comment (lib , "TKPrim.lib") 
#pragma comment (lib , "TKPShape.lib") 
#pragma comment (lib , "TKService.lib") 
#pragma comment (lib , "TKShHealing.lib") 
#pragma comment (lib , "TKTopAlgo.lib") 
#pragma comment (lib , "TKV3d.lib") 
//#pragma comment (lib , "TKV2d.lib")
#pragma comment (lib , "TKXSBase.lib") 

