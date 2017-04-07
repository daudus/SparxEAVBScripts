option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'


Dim SubP As EA.Package
Dim s As EA.Element
Set s = Nothing
Dim componentForGap
Dim ss
Dim concService 
Dim concFunction 
Dim concSystem 
Dim concInterface
Dim counter

sub main
	Dim AppLayerPackage As EA.Package
	Dim ReleasePackage As EA.Package
	Dim PBRPackage As EA.Package
	Dim Diagram As EA.Diagram
	Dim entries
	Set entries = CreateObject("Scripting.Dictionary")
	Set componentForGap = CreateObject("Scripting.Dictionary")
    Dim entr 
	Dim selectedObjectType
    Dim c
	Dim i
	Dim q
	Dim h
    c = 0
	counter = 0
	concService = "Service"
	concFunction = "Function"
	concSystem = "System"
	concInterface = "Interface"
	
	' just for testing purposes for particular package
    'Set PBRPackage = Repository.GetPackageByGuid("{B06008DE-7438-4be2-92BA-389333B25FDA}")
	
	' get package for evaluation
	Set ReleasePackage = Repository.GetContextObject() 
	selectedObjectType = Repository.GetContextItemType()
	if selectedObjectType <> otPackage then
		call Session.Prompt ("You have to select release package", 0)
		exit sub
	end if
     
	For c = 0 To ReleasePackage.Packages.count - 1 
		Set PBRPackage = ReleasePackage.Packages.GetAt(c)
		if InStr(1, PBRPackage.Name, "PBR") then
			' call Session.Output("PBR")
			For q = 0 To PBRPackage.Packages.count - 1 
				Set AppLayerPackage = PBRPackage.Packages.GetAt(q)
				if InStr(1, AppLayerPackage.Name,"App") then
					'call Session.Output(AppLayerPackage.Name)
					For i = 0 To AppLayerPackage.Diagrams.count - 1
						Set Diagram = AppLayerPackage.Diagrams.GetAt(i)
						 'call Session.Output("Diagram")
						For h = 0 To Diagram.DiagramObjects.count - 1
							Dim diagramObj As EA.DiagramObject
							Set diagramObj = Diagram.DiagramObjects(h)
							Dim obj As EA.Element
							Set obj = repository.GetElementByID(diagramObj.ElementID)
							If (obj.Stereotype = "ArchiMate_Gap") Then ' check only GAP elements
								
								Set entries = GetSourceElementsForGap(obj, PBRPackage.Name, entries)
							End If
						Next
					next
				end if
			next
		end if
	next
	
	' WriteOutput entries
end sub


Public Function WriteOutput(pEntries)
	 Dim e
	 Dim c 
	 call Session.Output ("PBR" + Chr(9) + "ServiceName" + Chr(9) + "GAP")
	 call Session.Output(pEntries.Count)
	 c = pEntries.Items
	 for e = 0 to pEntries.Count - 1
		Dim concept 
		Dim name
		if c(e).Concept = concService Then 
			name = c(e).AppService
		elseif c(e).Concept = concFunction Then 
			name = c(e).AppFunction
		elseif c(e).Concept = concSystem Then 
			name = c(e).AppComponent
		elseif c(e).Concept = concInterface Then 
			name = c(e).AppInterface
		end if
		
		call Session.Output(c(e).pbr + Chr(9) + name + Chr(9) + c(e).AppGap)
		
		' call Session.Output (c(e).pbr + ";" + c(e).AppGap +";"+ c(e).AppService +";"+ c(e).AppFunction +";"  +c(e).AppComponent)
	 Next
	 
End Function

Public Function GetSourceElementsForGap(gap , pbr , entries ) 
	 ' call Session.Output(pbr)
    dim connector as EA.Connector
    For Each connector In gap.Connectors
		Dim entr
        Set entr = New Entry
		counter = counter + 1
        If connector.Stereotype <> "ArchiMate_Association" Then
			
			 
			entries.add counter, entr
        Else
             Dim el As EA.Element
             
             ' --------- setting GAP basic attributes -----------
			 entr.AppGap = gap.Name
			 entr.Impact = ""
			 entr.GapDescription = ""
			 entr.Id = gap.ElementID
             entr.pbr = pbr
			 
			 ' --------- adding GAP to the output ---------------
			 
             entries.Add counter, entr
             If connector.ClientID = gap.ElementID Then
                    Set el = repository.GetElementByID(connector.SupplierID)
                    If (el.Stereotype = "ArchiMate_ApplicationService") Then
                        'call Session.Output("Service")
						entr.AppService = el.Name
						entr.Concept = concService
                        
                        
                    End If
                   
            End If
            If connector.SupplierID = gap.ElementID Then
                     Set el = repository.GetElementByID(connector.ClientID)
                     If (el.Stereotype = "ArchiMate_ApplicationService") Then
                        entr.AppService = el.Name
						entr.Concept = concService
                        
						
                    End If
                   
            End If
        End If   
		if not entr Is Nothing  Then
			if Len(entr.AppService) > 0 then
				call Session.Output(entr.pbr + Chr(9) + entr.AppService + Chr(9) + entr.AppGap)
			end if
		else 
			call Session.Output("!!!!!!!!!!!!" + gap.Name)
		end if
    Next 
	
    Set GetSourceElementsForGap = entries
End Function




class Entry
Private pPBR 
Private pAppFunction 
Private pAppService 
Private pAppComponent 
Private pAppGap
Private pImpact
Private pGapDescription
Private pConcept
Private pId 
Private pSourceSystem 
Public Property Get pbr() 
    pbr = pPBR
End Property
Public Property Let pbr(Value )
    pPBR = Value
End Property
Public Property Get AppFunction() 
    AppFunction = pAppFunction
End Property
Public Property Let AppFunction(Value)
    pAppFunction = Value
End Property
Public Property Get AppService() 
    AppService = pAppService
End Property
Public Property Let AppService(Value )
    pAppService = Value
End Property
Public Property Get AppGap() 
    AppGap = pAppGap
End Property
Public Property Let AppGap(Value )
    pAppGap = Value
End Property
Public Property Get Id()
    Id = pId
End Property
Public Property Let Id(Value )
    pId = Value
End Property
Public Property Get SourceSystem() 
    SourceSystem = pSourceSystem
End Property
Public Property Let SourceSystem(Value )
    pSourceSystem = Value
End Property
Public Property Get AppComponent() 
    AppComponent = pAppComponent
End Property
Public Property Let AppComponent(Value)
    pAppComponent = Value
End Property
Public Property Get Concept() 
    Concept = pConcept
End Property
Public Property Let Concept(Value)
    pConcept = Value
End Property
Public Property Get Impact() 
    Impact = pImpact
End Property
Public Property Let Impact(Value)
    pImpact = Value
End Property
Public Property Get GapDescription() 
    GapDescription = pGapDescription
End Property
Public Property Let GapDescription(Value)
    pGapDescription = Value
End Property
end class

main 

