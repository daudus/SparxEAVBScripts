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

sub main
	Dim AppLayerPackage As EA.Package
	Dim PBRPackage As EA.Package
	Dim entries
	Set entries = CreateObject("Scripting.Dictionary")
	Set componentForGap = CreateObject("Scripting.Dictionary")
    Dim entr 
	Dim selectedObjectType
    Dim c
    c = 0
	concService = "Service"
	concFunction = "Function"
	concSystem = "System"
	concInterface = "Interface"
	
	' just for testing purposes for particular package
    'Set PBRPackage = Repository.GetPackageByGuid("{B06008DE-7438-4be2-92BA-389333B25FDA}")
	
	' get package for evaluation
	Set PBRPackage = Repository.GetContextObject() 
	selectedObjectType = Repository.GetContextItemType()
	if selectedObjectType <> otPackage then
		call Session.Prompt ("You have to select requirement package", 0)
		exit sub
	end if
     
	' need to get to app layer package
	Set AppLayerPackage = PBRPackage.Packages.GetByName("Application Layer")
    If AppLayerPackage Is Nothing Then
           Set AppLayerPackage = PBRPackage.Packages.GetByName("Application layer")
    End If
	
	if AppLayerPackage is Nothing then
		call Session.Prompt ("Application L/layer package cannot be found - inappropriate package structure", 0)
		exit sub
	end if
                    
    Dim Diagram As EA.Diagram
    Dim i 
	If Not (AppLayerPackage.Diagrams.count = 0) Then
		For i = 0 To AppLayerPackage.Diagrams.count - 1
			Set Diagram = AppLayerPackage.Diagrams.GetAt(i)
				if Diagram.Name = "SA" Or  Diagram.Name ="Application Layer" Then ' iterate only through elements in diagram called "SA"
                    For c = 0 To Diagram.DiagramObjects.count - 1
						Dim diagramObj As EA.DiagramObject
                        Set diagramObj = Diagram.DiagramObjects(c)
                        Dim obj As EA.Element
						Set obj = repository.GetElementByID(diagramObj.ElementID)
					
						If (obj.Stereotype = "ArchiMate_Gap") Then ' check only GAP elements
							Set entries = GetSourceElementsForGap(obj, PBRPackage.Name, entries)
                        End If
                    Next
				End If
        Next
	else
		call Session.Prompt ("There's not diagram under application layer - inappropriate package structure", 0)
		exit sub
	end if
  WriteOutput entries
End Sub

Public Function GetSourceElementsForGap(gap , pbr , entries ) 
   

    dim connector as EA.Connector
    For Each connector In gap.Connectors
        If connector.Stereotype <> "ArchiMate_Association" Then
            exit function
        Else
             Dim el As EA.Element
             Dim entr
             Set entr = New Entry
             ' --------- setting GAP basic attributes -----------
			 entr.AppGap = gap.Name
			 entr.Impact = GetGapImpact(gap)
			 entr.GapDescription = GetGapDescription(gap)
			 entr.Id = gap.ElementID
             entr.pbr = pbr
			 
			 ' --------- adding GAP to the output ---------------
             entries.Add gap.ElementID, entr
             If connector.ClientID = gap.ElementID Then
                    Set el = repository.GetElementByID(connector.SupplierID)
                    If (el.Stereotype = "ArchiMate_ApplicationInterface") Then
                        entr.AppInterface= el.Name
						entr.Concept = concInterface
                        Dim temp As EA.Element
                        if isobject(GetSourceForAppInterface(el, entr)) then
							Set temp = GetSourceForAppInterface(el, entr)
						end if
                    End If
            End If
            If connector.SupplierID = gap.ElementID Then
                     Set el = repository.GetElementByID(connector.ClientID)
                     If (el.Stereotype = "ArchiMate_ApplicationInterface") Then
                        entr.AppInterface = el.Name
						entr.Concept = concInterface
                        Dim tempS As EA.Element
						if isobject(GetSourceForAppInterface(el, entr)) then
							Set tempS = GetSourceForAppInterface(el, entr)
						end if
                    End If
                 
            End If
        End If   
    Next 
    Set GetSourceElementsForGap = entries
End Function

Public Function GetGapDescription(gap)
	if not gap is Nothing Then
		GetGapDescription = gap.Notes
	end if 
End function

Public Function GetGapImpact(gap)
	Dim firstColonPositionIndex
	if not gap is Nothing and Len(gap.Name) > 0 Then
		firstColonPositionIndex = InStr(gap.Name, ":")
		if firstColonPositionIndex = 0 then
			call Session.Prompt ("Gap <" + gap.Name + "> does not have impact or character "":"" is missing ", 0)
			exit function
		end if
		GetGapImpact = Mid(gap.Name, 1, firstColonPositionIndex-1)
	end if
End function

Public Function GetSourceForAppInterface(inter, ByRef entr ) 
  
    Dim sourceFunction As EA.Element
    Dim sourceComponent As EA.Element
    Dim cn 
    ' call Session.Output ("verification of service" + ser.Name)
    if inter.Name = "EXT" Then
		For Each cn In ser.Connectors
			Dim el As EA.Element
			If cn.Stereotype = "ArchiMate_Association" Then
				If cn.SupplierID = inter.ElementID Then
					  Set el = repository.GetElementByID(cn.ClientID)
					  call Session.Output ("app service" + inter.Name)
				End If
			   
			End If
		Next 
	end if 
    entr.AppService = ser.Name
	If isobject(sourceComponent)Then
        Set GetSourceForAppService = sourceComponent
    End If
	
End Function

class Entry
Private pPBR 
Private pAppFunction 
Private pAppService 
Private pAppInterface
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

Public Property Get AppInterface() 
    AppInterface = pAppInterface
End Property
Public Property Let AppInterface(Value )
    pAppInterface = Value
End Property

end class

main 
