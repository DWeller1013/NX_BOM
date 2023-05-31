'--------------------------------------------------------------------------------------------
'	230330(dmw) - Adjusted if statement in printAttribute function to also check parentName
'				  when incremeneting. It should seperate detail #'s correctly now based on
'				  USH vs LSH...
'--------------------------------------------------------------------------------------------
'	230213(dmw) - Added a new property to part class to get parent name. Adjusted excel sheet
'				  to add a new column "Owning Comp" to include parent name. 
'--------------------------------------------------------------------------------------------
'	221114(dmw) - Program to create a BOM excel file from template provided.
'				  If the program is ran again while a BOM already exists, it creates a 
'				  new sheet within the same BOM file.
'				  BOM pulls specific attributes from all the parts within the assembly.
'--------------------------------------------------------------------------------------------
Option Strict Off
Imports System
Imports System.IO
Imports System.Environment
Imports System.Collections
Imports NXOpen
Imports NXOpen.UF
Imports NXOpen.Assemblies
Imports System.Collections.Generic
'---------------------------------------------------------------------------------------------
Module NXJournal
'---------------------------------------------------------------------------------------------  
    ' Declare public variables used across entire module (All Subs / Functions)
	Public theSession As Session = Session.GetSession()
    Public theUFSession As UFSession = UFSession.GetUFSession()
    Dim attributeList As New List(Of AssemblyComponent)
    Dim fileName As String
    Dim filePath As String
    Dim directoryPath As String
    Dim templatePath As String = "\\server\files\lib\macros\NX\TEMPLATES\STOCKLIST-STARTER.xlsm" ' Change path to where template is stored.
'---------------------------------------------------------------------------------------------    
    Sub Main (ByVal args As String())
    
		' Declare work part
        Dim workPart As Part = theSession.Parts.Work
       
	    ' Check if workPart is empty
		If workPart Is Nothing Then

			If args.Length = 0 Then
				Echo("Part file argument expected or work part required")
				Return
			End If

			theSession.Parts.LoadOptions.ComponentsToLoad = _
				LoadOptions.LoadComponents.None
			theSession.Parts.LoadOptions.SetInterpartData(True, _
				LoadOptions.Parent.All)
			theSession.Parts.LoadOptions.UsePartialLoading = True

			Dim partLoadStatus1 As PartLoadStatus = Nothing
			workPart = theSession.Parts.OpenBaseDisplay(args(0), partLoadStatus1)

			If workPart Is Nothing Then Return
		End If
		
		' Check if c.RootComponent exists and print main assembly name and call bomCategories
        Dim c As ComponentAssembly = workPart.ComponentAssembly
        
        If not IsNothing(c.RootComponent) Then
            
            fileName = c.RootComponent.DisplayName
			filePath = workPart.FullPath
			directoryPath = Path.GetDirectoryName(filePath)
            bomCategories(c.RootComponent)
            
        End If
        
        'AttributeList.RemoveAll(Function(rmAttr) AssemblyComponent.DetailNumber = "")
        cleanList()
        'createExcel()
        excelTemplate()
        
    End Sub
'---------------------------------------------------------------------------------------------
   ' Sub routine to remove all objects in list that do not contain a detail number attribute. 
   Sub cleanList()
        
        'Guide.InfoWriteLine(attributeList))
        'For Each i As AssemblyComponent In attributeList
        For i As Integer = attributeList.Count - 1 to 0 Step -1
            If attributeList(i).getDetailNumber() = " " or attributeList(i).getDetailNumber() = "" Then
                'Guide.InfoWriteLine("Test i: " & attributeList(i).getDetailNumber())
                attributeList.RemoveAt(i)
            End If 
        Next
        
    End Sub
'---------------------------------------------------------------------------------------------
    Sub bomCategories(comp As NXOpen.Assemblies.Component)
    
        Dim attributes = theSession.Parts.Work.GetUserAttributes()
        
        For Each child As Component In comp.GetChildren()

            printAttributes(child, child)
            bomCategories(child)
        
            Dim markId1 As NXOpen.Session.UndoMarkId = Nothing
            markId1 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Make Work Part")
            
            Dim theUI As UI = UI.GetUI()
        
        Next
        
    End Sub
'---------------------------------------------------------------------------------------------
    ' Sub routine to add objects to attribute list. If object already exists in lists, increment the quantity instead.
	Sub printAttributes(obj As NXObject, objComp As Component)
        
        Dim attributes = obj.GetUserAttributes()
        Dim partStr As String
        partStr = Convert.ToString(obj.Name)
        
        'If (partStr.Contains("_D") or partStr.Contains("_d") or partStr.Contains("-D") or partStr.Contains("-d")) Then
      
        Dim partName As AssemblyComponent = New AssemblyComponent(partStr)

        ' 230330(dmw) - TESTING
        'Guide.InfoWriteLine("parentName: " & assyPart.getParentName())
        'Guide.InfoWriteLine("objComp: " & objComp.Parent.Name)

        For Each assyPart As AssemblyComponent In AttributeList 
        
            'Guide.InfoWriteLine("parentName: " & assyPart.getParentName())
            'Guide.InfoWriteLine("partName: " & partName.getPartName())
            'Guide.InfoWriteLine("assyPart: " & assyPart.getPartName())
            
			' 230330(dmw) - Add 'And' in If statement to also check parent name.
			' If part exists, exit out of sub and increment the quantity
            If (partName.getPartName() = assyPart.getPartName() And assyPart.getParentName() = objComp.Parent.Name) Then
                
                assyPart.incrementQty()
                'Guide.InfoWriteLine(assyPart.getQty())

                Return
                
            End If
            
        Next assyPart
		
		' Loop through each attribute for part and assign to a property
        For Each attribute As NXObject.AttributeInformation In attributes
    
            If (attribute.Title = "Change") Then
                partName.setChange(attribute.StringValue)
            End If
            
            If (attribute.Title = "Comment") Then
                partName.setComment(attribute.StringValue)
            End If
            
            If (attribute.Title = "Description") Then
                partName.setDescription(attribute.StringValue)
            End If                
            
            If (attribute.Title = "Detail Number") Then
                partName.setDetailNumber(attribute.StringValue)
            End If   
            
            If (attribute.Title = "Heat_Treat") Then
                partName.setHeatTreat(attribute.StringValue)
            End If   
            
            If (attribute.Title = "Dimension") Then
                partName.setDimension(attribute.StringValue)
            End If   
            
            If (attribute.Title = "Material/Ordering Number") Then
                partName.setMaterial(attribute.StringValue)
            End If  
            
        Next attribute

		'----------------------------------------------------------------------	
		' 230213(dmw) - Test adding parent name as partName property. 
		'----------------------------------------------------------------------
		'Guide.InfoWriteLine(objComp.Parent.Name) '.Name)
		If (objComp.Parent.Name = "" OR objComp.Parent.Name = " ") Then
			partName.setParentName("ENTIRE DESIGN")
		Else
			partName.setParentName(objComp.Parent.Name)
		End If
		'Guide.InfoWriteLine(partName.getParentName)
		'-----------------------------------------------------------------------        
        
        attributeList.Add(partName)
  
        'End If
        
        ' TEST
        'Guide.InfoWriteLine("Test part: ")
        'Guide.InfoWriteLine(partName.getChange)
        'Guide.InfoWriteLine(partName.getComment)
        'Guide.InfoWriteLine(partName.getDescription)
        'Guide.InfoWriteLine(partName.getDetailNumber)
        'Guide.InfoWriteLine(partName.getHeatTreat)
        'Guide.InfoWriteLine(partName.getDimension)
        'Guide.InfoWriteLine(partName.getMaterial)
        'Guide.InfoWriteLine(partName.getParentName)
        'Guide.InfoWriteLine(" ")
        ' END TEST

    End Sub
'---------------------------------------------------------------------------------------------
    ' Sub routine to create excel file and add headers / all attributeList objects to cells.
	' Saves to C:\tmp\
	Sub createExcel()
    
        Dim oExcel As Object    
        Dim oBook As Object    
        Dim oSheet As Object 
            
        'Start a new workbook in Excel    
        oExcel = CreateObject("Excel.Application")    
        oBook = oExcel.Workbooks.Add
              
        'Add data to cells of the first worksheet in the new workbook    
        oSheet = oBook.Worksheets(1) 
        oSheet.Name = fileName
        'oSheet.Range("A1").Value = "Part Name" 
        oSheet.Range("A1").Value = "Detail Number"
        oSheet.Range("B1").Value = "Quantity"    
        oSheet.Range("C1").Value = "Description"
        oSheet.Range("D1").Value = "Material/Ordering Number"
        oSheet.Range("E1").Value = "Dimension"
        oSheet.Range("F1").Value = "Heat_Treat"
        oSheet.Range("G1").Value = "Comment"
        oSheet.Range("H1").Value = "Change"
		oSheet.Range("I1").Value = "Owning Comp"
        oSheet.Range("A1:I1").Font.Bold = True 
        
        Dim excelRow As Integer
        excelRow = 2
        
        ' Sort list by detail num from low to high
        attributeList.Sort(Function(x, y) x.getDetailNumber.CompareTo(y.getDetailNumber))
        
        For Each assyPart As AssemblyComponent In attributeList
        
            'oSheet.Range("A" & excelRow).Value = assyPart.getPartName()
            oSheet.Range("A" & excelRow).Value = assyPart.getDetailNumber()
            oSheet.Range("B" & excelRow).Value = assyPart.getQty()
            oSheet.Range("C" & excelRow).Value = assyPart.getDescription()
            oSheet.Range("D" & excelRow).Value = assyPart.getMaterial()
            oSheet.Range("E" & excelRow).Value = assyPart.getDimension()
            oSheet.Range("F" & excelRow).Value = assyPart.getHeatTreat()
            oSheet.Range("G" & excelRow).Value = assyPart.getComment()
            oSheet.Range("H" & excelRow).Value = assyPart.getChange()
			oSheet.Range("I" & excelRow).Value = assyPart.getParentName()
            excelRow = excelRow + 1
            
        Next
        
        ' -4108 is the value for xlCenter. Center align both vertically and horizontally for all columns / rows
        oSheet.Range("A1:I" & excelRow).HorizontalAlignment = -4108
        oSheet.Range("A1:I" & excelRow).VerticalAlignment = -4108
        oSheet.Range("A1:I" & excelRow).Columns.AutoFit
        
        ' Grab current date and Save BOM with date on the end.
        Dim currDate As Date = Now
        Dim strDate As String = currDate.ToString("MM-dd-yy")
        
        'Save the Workbook and Quit Excel    
        'oBook.SaveAs("C:\tmp\" & fileName & "-BOM_" & strDate & ".xlsx")    
        'Guide.InfoWriteLine(directoryPath)
        Dim fullPath As String = directoryPath & "\" & fileName & "-BOM" & ".xlsx"
        oBook.SaveAs(fullPath) 
        oExcel.Quit
        
    End Sub

'---------------------------------------------------------------------------------------------
	Sub excelTemplate()
	
        Dim oExcel As Object    
        Dim oBook As Object    
        Dim oSheet As Object 
        Dim counter As Integer = 1
        Dim currDate As Date = Now
        Dim strDate As String = currDate.ToString("MM-dd-yy")
        Dim fullPath As String = directoryPath & "\" & fileName & "-BOM" & ".xlsm"
        'Dim sht As Excel.Worksheet
         
        'Start a new workbook in Excel    
        oExcel = CreateObject("Excel.Application")    
        
        ' Check if file exist in job folder. If it doesn't exist, use template
		If System.IO.File.Exists(fullPath) Then
			oBook = oExcel.Workbooks.Open(fullPath)
            For i As Integer = 1 To oBook.WorkSheets.Count
                If InStr(oBook.Worksheets(i).Name, "BOM") Then
                    'Guide.InfoWriteLine("i: " & i)
                    counter = counter + 1
                End If 
            Next i
		Else
			oBook = oExcel.Workbooks.Open(templatePath)
			counter = 1
		End If
           
        oBook.Worksheets.Add
        oSheet = oBook.Worksheets("Sheet1")
        
        If (counter > 9) Then 
            oSheet.Name = "BOM-0" & counter
        Else 
            oSheet.Name = "BOM-00" & counter
        End If 
        
        Dim shtLength
        shtLength = oBook.Worksheets.Count
        oSheet.Move(Before: = oBook.Worksheets(shtLength))
        
        oSheet.Range("A1").Value = "CAD Model Name:"
        oSheet.Range("A2").Value = "Die No:"
        oSheet.Range("A3").Value = "Die Description:"
        oSheet.Range("A4").Value = "Job No:"
        oSheet.Range("A5").Value = "Die Change Level:"
        oSheet.Range("A1:A5").Font.Bold = True
        
        oSheet.Range("A7").Value = "Detail Number"
        oSheet.Range("B7").Value = "Quantity"    
        oSheet.Range("C7").Value = "Description"
        oSheet.Range("D7").Value = "Material/Ordering Number"
        oSheet.Range("E7").Value = "Dimension"
        oSheet.Range("F7").Value = "Heat_Treat"
        oSheet.Range("G7").Value = "Comment"
        oSheet.Range("H7").Value = "Change"
		oSheet.Range("I7").Value = "Owning Comp"
        oSheet.Range("A7:I7").Font.Bold = True 
		oSheet.Range("A7:I7").AutoFilter ' 230320(dmw) - Adding AutoFilter toggle to always be on.       
 
        oSheet.Range("B1").Formula = "='INFO PAGE'!B4"
        oSheet.Range("B2").Formula = "='INFO PAGE'!B5"
        oSheet.Range("B3").Formula = "='INFO PAGE'!B6"
        oSheet.Range("B4").Formula = "='INFO PAGE'!B7"
        oSheet.Range("B5").Formula = "='INFO PAGE'!B8"
        
        Dim excelRow As Integer
        excelRow = 8
        
        ' Sort list by detail num from low to high
        attributeList.Sort(Function(x, y) x.getDetailNumber.CompareTo(y.getDetailNumber))

		' 230406(dmw)-Create object and qty to set for current detail
		' Initialize variables and set to empty strings / blank qty
		Dim currPart As AssemblyComponent = New AssemblyComponent("currPart")
		Dim qty As Integer       
		qty = 0
        currPart.setQty(qty)
        
		For Each assyPart As AssemblyComponent In attributeList
        
        'Guide.InfoWriteLine("This detail#: " & assyPart.getDetailNumber())
        'Guide.InfoWriteLine("Current detail#: " & currPart.getDetailNumber())
        'Guide.InfoWriteLine("")
        
			' 230406(dmw)-Add currPart / qty, if / else statements.
			If currPart.getDetailNumber() = assyPart.getDetailNumber() Then
				
				qty = qty + assyPart.getQty()
				currPart.setQty(qty)	
				
				'Guide.InfoWriteLine("")
				'Guide.InfoWriteLine("Detail match...")
				'Guide.InfoWriteLine("total qty: " & currPart.getQty())
				'Guide.InfoWriteLine("curr part qty: " & assyPart.getQty())
			
			Else
			
				 If qty = 0 Then	
				'If currPart.getQty() = 0 Then				
			
					currPart.setDetailNumber(assyPart.getDetailNumber())
					'currPart.setQty(0)
					qty = assyPart.getQty()
					currPart.setQty(assyPart.getQty())
					
					'Guide.InfoWriteLine("")
					'Guide.InfoWriteLine("Initial quantity: ")
					'Guide.InfoWriteLine("total qty: " & currPart.getQty())
					'Guide.InfoWriteLine("curr part qty: " & assyPart.getQty())
					
					
					'currPart.setQty(qty)
					currPart.setDescription(assyPart.getDescription())
					currPart.setMaterial(assyPart.getMaterial())	
					currPart.setDimension(assyPart.getDimension())
					currPart.setComment(assyPart.getComment())
					currPart.setHeatTreat(assyPart.getHeatTreat())
					currPart.setChange(assyPart.getChange())
					currPart.setParentName("ENTIRE DESIGN")
					
				Else
                    
                    'Guide.InfoWriteLine("")
					'Guide.InfoWriteLine("Detail change... ")
					
					oSheet.Range("A" & excelRow).Value = currPart.getDetailNumber()
            		oSheet.Range("B" & excelRow).Value = currPart.getQty()
            		oSheet.Range("C" & excelRow).Value = currPart.getDescription()
            		oSheet.Range("D" & excelRow).Value = currPart.getMaterial()
            		oSheet.Range("E" & excelRow).Value = currPart.getDimension()
            		oSheet.Range("F" & excelRow).Value = currPart.getHeatTreat()
            		oSheet.Range("G" & excelRow).Value = currPart.getComment()
            		oSheet.Range("H" & excelRow).Value = currPart.getChange()
            		oSheet.Range("I" & excelRow).Value = currPart.getParentName()
            		excelRow = excelRow + 1
            		qty = assyPart.getQty()
            		currPart.setDetailNumber(assyPart.getDetailNumber())
					currPart.setDescription(assyPart.getDescription())
					currPart.setMaterial(assyPart.getMaterial())	
					currPart.setDimension(assyPart.getDimension())
					currPart.setComment(assyPart.getComment())
					currPart.setHeatTreat(assyPart.getHeatTreat())
					currPart.setChange(assyPart.getChange())
					currPart.setParentName("ENTIRE DESIGN")
					currPart.setQty(qty)
            		'currPart = New AssemblyComponent("currPart")
				    '
				    
				End If 
	
			End If 

			oSheet.Range("A" & excelRow).Value = assyPart.getDetailNumber()
            oSheet.Range("B" & excelRow).Value = assyPart.getQty()
            oSheet.Range("C" & excelRow).Value = assyPart.getDescription()
            oSheet.Range("D" & excelRow).Value = assyPart.getMaterial()
            oSheet.Range("E" & excelRow).Value = assyPart.getDimension()
            oSheet.Range("F" & excelRow).Value = assyPart.getHeatTreat()
            oSheet.Range("G" & excelRow).Value = assyPart.getComment()
            oSheet.Range("H" & excelRow).Value = assyPart.getChange()
            oSheet.Range("I" & excelRow).Value = assyPart.getParentName()
            excelRow = excelRow + 1

        Next
        
        ' 230406(dmw) - Print the final currPart in last excel row.
        ' Unless qty = 0, then do nothing.
        If (currPart.getQty() > 0) Then
            oSheet.Range("A" & excelRow).Value = currPart.getDetailNumber()
            oSheet.Range("B" & excelRow).Value = currPart.getQty()
            oSheet.Range("C" & excelRow).Value = currPart.getDescription()
            oSheet.Range("D" & excelRow).Value = currPart.getMaterial()
            oSheet.Range("E" & excelRow).Value = currPart.getDimension()
            oSheet.Range("F" & excelRow).Value = currPart.getHeatTreat()
            oSheet.Range("G" & excelRow).Value = currPart.getComment()
            oSheet.Range("H" & excelRow).Value = currPart.getChange()
            oSheet.Range("I" & excelRow).Value = currPart.getParentName()
        End If 

        ' -4108 is the value for xlCenter. Center align both vertically and horizontally for all columns / rows
        oSheet.Range("A7:I" & excelRow).HorizontalAlignment = -4108
        oSheet.Range("A7:I" & excelRow).VerticalAlignment = -4108
        oSheet.Range("A1:I" & excelRow).Columns.AutoFit
        
        'oSheet.TabColor = #519FF5
        'Color.Blue
        
        oBook.Worksheets("INFO PAGE").Activate
        oBook.Worksheets("INFO PAGE").Range("A1").Select
        
        oBook.SaveAs(fullPath) 
        oExcel.Quit

	End Sub
'---------------------------------------------------------------------------------------------
	Function GetComponentFullPath(ByVal comp As Assemblies.Component) As String
		Dim partName As String = ""
		Dim refsetName As String = ""
		Dim instanceName As String = ""
		Dim origin(2) As Double
		Dim csysMatrix(8) As Double
		Dim transform(3, 3) As Double
		theUFSession.Assem.AskComponentData(comp.Tag, partName, refsetName, _
		 instanceName, origin, csysMatrix, transform)

		Return partName
	End Function
'---------------------------------------------------------------------------------------------
	Sub GetComponentsFullPaths(ByVal thisComp As Assemblies.Component, _
		ByVal parts As ArrayList)
		
		Dim thisPath As String = GetComponentFullPath(thisComp)
		If Not parts.Contains(thisPath) Then parts.Add(thisPath)
		For Each child As Assemblies.Component In thisComp.GetChildren()
			GetComponentsFullPaths(child, parts)
		Next
	End Sub
'---------------------------------------------------------------------------------------------
	Sub Echo(ByVal output As String)
		theSession.ListingWindow.Open()
		theSession.ListingWindow.WriteLine(output)
		theSession.LogFile.WriteLine(output)
	End Sub
'---------------------------------------------------------------------------------------------
    ' Component object to store each part found in assembly navigator
    Public Class AssemblyComponent
    
        Private quantity As Integer
        Private partName As String
        Private change As String
        Private comment As String
        Private description As String
        Private detailNumber As String
        Private dimension As String
        Private heatTreat As String
        Private material As String
		Private parentName As String
        
        Public Sub New(ByVal prtName As String)
            quantity = 1
            partName = prtName
            change = ""
            comment = ""
            description = ""
            detailNumber = ""
            dimension = ""
            heatTreat = ""
            material = ""
			parentName = ""
        End Sub
        
        Public Sub incrementQty()
            quantity = quantity + 1
        End Sub
                
        ' Getters and setters
        Public Function getPartName() As String
            Return partName
        End Function
        
        Public Function setQty(ByVal qtyVal As Integer) 
            quantity = qtyVal
        End Function
    
        Public Function getQty() As Integer
            Return quantity
        End Function

        Public Sub setChange(ByVal changeVal As String)
            change = changeVal
        End Sub
        
        Public Function getChange() As String
            Return change
        End Function
        
        Public Sub setComment(ByVal commentVal As String)
            comment = commentVal
        End Sub
        
        Public Function getComment() As String
            Return comment
        End Function
        
        Public Sub setDescription(ByVal descriptionVal As String)
            description = descriptionVal
        End Sub
        
        Public Function getDescription() As String
            Return description
        End Function
        
        Public Sub setDetailNumber(ByVal detailNumberVal As String)
            detailNumber = detailNumberVal
        End Sub
        
        Public Function getDetailNumber() As String
            Return detailNumber
        End Function
        
        Public Sub setDimension(ByVal dimensionVal As String)
            dimension = dimensionVal
        End Sub
        
        Public Function getDimension() As String
            Return dimension
        End Function
        
        Public Sub setHeatTreat(ByVal heatTreatVal As String)
            heatTreat = heatTreatVal
        End Sub
        
        Public Function getHeatTreat() As String
            Return heatTreat
        End Function
        
        Public Sub setMaterial(ByVal materialVal As String)
            material = materialVal
        End Sub
        
        Public Function getMaterial() As String
            Return material
        End Function
       
		Public Sub setParentName(ByVal parentNameVal As String)
			parentName = parentNameVal
		End Sub

		Public Function getParentName() As String
			Return parentName
		End Function
 
    End Class
'---------------------------------------------------------------------------------------------
End Module
