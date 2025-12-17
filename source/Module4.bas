Attribute VB_Name = "Module4"
Option Explicit
'// COPYRIGHT DASSAULT SYSTEMES  2000
'******************************************************************************
' Purpose:       This CATScript demonstrates how to create an ArrangementBoundary
'                and change it's visualization to "Solid" mode, define a
'                Rectangular section data and apply a constant bend radius of 25 mm.
' Assumptions:   This assumes that a macro is being executed interactively.
' Author     :
' Languages  :   VBScript
' CATIA Level:   V5R6
' Locale     :   English
'******************************************************************************


Sub CATMain()
   
   ' On Error Resume Next

   '----------------------------------------------
   'Create a new product document
   Dim objProdDoc        As ProductDocument
   Dim objRootProd       As Product
   Set objProdDoc = CATIA.Documents.Add("Product")
   Set objRootProd = objProdDoc.Product

   '----------------------------------------------
   'Retrieving Root Product's Relative Axis and Position Information
   Dim objMove           As Move
   Set objMove = objRootProd.Move

   '----------------------------------------------
   ' Get ArrangementProduct
   Dim objArrProd        As ArrangementProduct
   Set objArrProd = objRootProd.GetTechnologicalObject("ArrangementProduct")

   '----------------------------------------------
   ' Create ArrangementBoundary under the Root Product
   Dim dblBoundaryPoints(75)      As Double
   Dim dblMathDirection(3)        As Double
   Dim objArrBoundary             As ArrangementBoundary


   dblBoundaryPoints(0) = 300#
   dblBoundaryPoints(1) = 100#
   dblBoundaryPoints(2) = 0#

   dblBoundaryPoints(3) = 441.42
   dblBoundaryPoints(4) = 158.58
   dblBoundaryPoints(5) = 1.25

   dblBoundaryPoints(6) = 500#
   dblBoundaryPoints(7) = 300#
   dblBoundaryPoints(8) = 2.5

   dblBoundaryPoints(9) = 441.42
   dblBoundaryPoints(10) = 441.42
   dblBoundaryPoints(11) = 3.75

   dblBoundaryPoints(12) = 300#
   dblBoundaryPoints(13) = 500#
   dblBoundaryPoints(14) = 5#

   dblBoundaryPoints(15) = 158.58
   dblBoundaryPoints(16) = 441.42
   dblBoundaryPoints(17) = 6.25

   dblBoundaryPoints(18) = 100#
   dblBoundaryPoints(19) = 300#
   dblBoundaryPoints(20) = 7.5

   dblBoundaryPoints(21) = 158.58
   dblBoundaryPoints(22) = 158.58
   dblBoundaryPoints(23) = 8.75

   dblBoundaryPoints(24) = 300#
   dblBoundaryPoints(25) = 100#
   dblBoundaryPoints(26) = 10

   dblBoundaryPoints(27) = 441.42
   dblBoundaryPoints(28) = 158.58
   dblBoundaryPoints(29) = 11.25

   dblBoundaryPoints(30) = 500#
   dblBoundaryPoints(31) = 300#
   dblBoundaryPoints(32) = 12.5

   dblBoundaryPoints(33) = 441.42
   dblBoundaryPoints(34) = 441.42
   dblBoundaryPoints(35) = 13.75

   dblBoundaryPoints(36) = 300#
   dblBoundaryPoints(37) = 500#
   dblBoundaryPoints(38) = 15#

   dblBoundaryPoints(39) = 158.58
   dblBoundaryPoints(40) = 441.42
   dblBoundaryPoints(41) = 16.25

   dblBoundaryPoints(42) = 100#
   dblBoundaryPoints(43) = 300#
   dblBoundaryPoints(44) = 17.5

   dblBoundaryPoints(45) = 158.58
   dblBoundaryPoints(46) = 158.58
   dblBoundaryPoints(47) = 18.75

   dblBoundaryPoints(48) = 300#
   dblBoundaryPoints(49) = 100#
   dblBoundaryPoints(50) = 20

   dblBoundaryPoints(51) = 441.42
   dblBoundaryPoints(52) = 158.58
   dblBoundaryPoints(53) = 21.25

   dblBoundaryPoints(54) = 500#
   dblBoundaryPoints(55) = 300#
   dblBoundaryPoints(56) = 22.5

   dblBoundaryPoints(57) = 441.42
   dblBoundaryPoints(58) = 441.42
   dblBoundaryPoints(59) = 23.75

   dblBoundaryPoints(60) = 300#
   dblBoundaryPoints(61) = 500#
   dblBoundaryPoints(62) = 25#

   dblBoundaryPoints(63) = 158.58
   dblBoundaryPoints(64) = 441.42
   dblBoundaryPoints(65) = 26.25

   dblBoundaryPoints(66) = 100#
   dblBoundaryPoints(67) = 300#
   dblBoundaryPoints(68) = 27.5

   dblBoundaryPoints(69) = 158.58
   dblBoundaryPoints(70) = 158.58
   dblBoundaryPoints(71) = 28.75

   dblBoundaryPoints(72) = 300#
   dblBoundaryPoints(73) = 100#
   dblBoundaryPoints(74) = 30


   dblMathDirection(0) = 1#
   dblMathDirection(1) = 0#
   dblMathDirection(2) = 0#

   Set objArrBoundary = objArrProd.ArrangementBoundaries.AddBoundary(objMove, dblBoundaryPoints, dblMathDirection)

   '----------------------------------------------
   ' Change Properties of ArrangementBoundary
   objArrBoundary.SectionType = CatArrangementRouteSectionRectangular
   objArrBoundary.SectionWidth = 10#
   objArrBoundary.SectionHeight = 10#
   objArrBoundary.VisuMode = CatArrangementRouteVisuModeSolid

   '----------------------------------------------
   ' Define Bend Radius of Nodes
   Dim intK  As Integer
   For intK = 1 To objArrBoundary.ArrangementNodes.count
     objArrBoundary.ArrangementNodes.item(intK).BendRadius = 10#
   Next


End Sub


