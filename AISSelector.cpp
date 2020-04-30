#include "StdAfx.h"
#include "AISSelector.h"

CAISSelector::CAISSelector(void)
{
}

CAISSelector::~CAISSelector(void)
{
}

// Copyright (C) 2007-2012  CEA/DEN, EDF R&D, OPEN CASCADE
//
// Copyright (C) 2003-2007  OPEN CASCADE, EADS/CCR, LIP6, CEA/DEN,
// CEDRAT, EDF R&D, LEG, PRINCIPIA R&D, BUREAU VERITAS
//
// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public
// License as published by the Free Software Foundation; either
// version 2.1 of the License.
//
// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
// Lesser General Public License for more details.
//
// You should have received a copy of the GNU Lesser General Public
// License along with this library; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307 USA
//
// See http://www.salome-platform.org/ or email : webmaster.salome@opencascade.com
//

//#include "OCCViewer_AISSelector.h"
//
//OCCViewer_AISSelector::OCCViewer_AISSelector( QObject* parent, 
//const Handle (AIS_InteractiveContext)& aisContext) :
//QObject( parent ), 
//myNumSelected( 0 ), 
//myEnableSelection( true ),
//myEnableMultipleSelection( true )
//{
//myHilightColor = Quantity_NOC_CYAN1;
//mySelectColor  = Quantity_NOC_GRAY80;
//
//setAISContext( aisContext );
//}
//
//OCCViewer_AISSelector::~OCCViewer_AISSelector()
//{
//}
//
//void OCCViewer_AISSelector::enableSelection( bool bEnable )
//{
//myEnableSelection = bEnable;
//}
//
//void OCCViewer_AISSelector::enableMultipleSelection( bool bEnable )
//{
//myEnableMultipleSelection = bEnable;
//if ( bEnable ) myEnableSelection = bEnable;
//}
//
//void OCCViewer_AISSelector::setHilightColor ( Quantity_NameOfColor color )
//{
//myHilightColor = color;
//if ( !myAISContext.IsNull() )
//myAISContext->SetHilightColor( myHilightColor );
//}
//
//void OCCViewer_AISSelector::setSelectColor ( Quantity_NameOfColor color )
//{
//mySelectColor = color;
//if ( !myAISContext.IsNull() )
//myAISContext->SelectionColor( mySelectColor );
//}
//
//void OCCViewer_AISSelector::setAISContext ( const Handle (AIS_InteractiveContext)& aisContext )
//{
//myAISContext = aisContext;
//if ( ! myAISContext.IsNull() ) { 
//myAISContext->SetHilightColor( myHilightColor );
//myAISContext->SelectionColor( mySelectColor );
//myAISContext->SetSubIntensityColor( Quantity_NOC_CYAN1 );
//}
//}
//
//bool OCCViewer_AISSelector::checkSelection ( AIS_StatusOfPick status, 
//bool hadSelection, 
//bool addTo )
//{
//if ( myAISContext.IsNull() )
//return false;
//
//myNumSelected = myAISContext->NbCurrents(); /* update after the last selection */
//
//if ( status == AIS_SOP_NothingSelected && !hadSelection ) {
//emit selSelectionCancel( addTo );
//}
//else if ( status == AIS_SOP_NothingSelected && hadSelection ) {
//emit selSelectionCancel( addTo ); /* unselected now */
//}
//else if ( status == AIS_SOP_OneSelected || status == AIS_SOP_SeveralSelected )
//{
//emit selSelectionDone( addTo ); /* selected ( the same object, may be ) */
//}
//return ( status != AIS_SOP_Error && status != AIS_SOP_NothingSelected );
//}
//
//bool OCCViewer_AISSelector::moveTo ( int x, int y, const Handle (V3d_View)& view )
//{
//if ( myAISContext.IsNull() )
//return false;
//
//if ( !myEnableSelection )
//return false;
//
//AIS_StatusOfDetection status = AIS_SOD_Error;
//status = myAISContext->MoveTo (x, y, view);
//
//return ( status != AIS_SOD_Error && status != AIS_SOD_AllBad );
//}
//
//bool OCCViewer_AISSelector::select ()
//{
//if ( myAISContext.IsNull() )
//return false;
//
//if ( !myEnableSelection )
//return false;
//
//bool hadSelection = ( myNumSelected > 0 );
//
///* select and send notifications */
//return checkSelection ( myAISContext->Select(), hadSelection, false );
//}
//
//bool OCCViewer_AISSelector::select ( int left, int top, int right, int bottom,
//const Handle (V3d_View)& view )
//{
//if ( myAISContext.IsNull() )
//return false;
//
//if ( !myEnableSelection || !myEnableMultipleSelection )
//return false;  /* selection with rectangle is considered as multiple selection */
//
//bool hadSelection = ( myNumSelected > 0 );
//
///* select and send notifications */
//return checkSelection ( myAISContext->Select(left, top, right, bottom, view),
//hadSelection, false );
//}
//
//bool OCCViewer_AISSelector::shiftSelect ()
//{
//if ( myAISContext.IsNull() )
//return false;
//
//if ( !myEnableSelection )
//return false;
//
//bool hadSelection = ( myNumSelected > 0 ); /* something was selected */
//if ( hadSelection && !myEnableMultipleSelection)
//return false;
//
///* select and send notifications */
//return checkSelection ( myAISContext->ShiftSelect(), hadSelection, true );
//}
//
//bool OCCViewer_AISSelector::shiftSelect ( int left, int top, int right, int bottom,
//const Handle (V3d_View)& view )
//
//{
//if ( myAISContext.IsNull() )
//return false;
//
//if ( !myEnableSelection || !myEnableMultipleSelection )
//return false;  /* selection with rectangle is considered as multiple selection */
//
//bool hadSelection = ( myNumSelected > 0 );      /* something was selected */
//if ( hadSelection && !myEnableMultipleSelection)
//return false;
//
///* select and send notifications */
//   return checkSelection ( myAISContext->ShiftSelect(left,top,right,bottom, view),
//     hadSelection, true );
//}