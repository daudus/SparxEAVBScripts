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
				if Diagram.Name = "SA" Then ' iterate only through elements in diagram called "SA"
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


Public Function WriteOutput(pEntries)
	 Dim e
	 Dim c 
	 call Session.Output ("Component" + Chr(9) + "Concept" + Chr(9) + "Name" + Chr(9) + "Impact" + Chr(9) + "Note")
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
		
		call Session.Output(c(e).AppComponent + Chr(9) + c(e).Concept + Chr(9) + name + Chr(9) + c(e).Impact + Chr(9) + c(e).GapDescription)
		
		' call Session.Output (c(e).pbr + ";" + c(e).AppGap +";"+ c(e).AppService +";"+ c(e).AppFunction +";"  +c(e).AppComponent)
	 Next
	 
End Function

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
                    If (el.Stereotype = "ArchiMate_ApplicationService") Then
                        entr.AppService = el.Name
						entr.Concept = concService
                        Dim temp As EA.Element
                        if isobject(GetSourceForAppService(el, entr)) then
							Set temp = GetSourceForAppService(el, entr)
						end if
                    End If
                    If (el.Stereotype = "ArchiMate_ApplicationFunction") Then
                        entr.AppFunction = el.Name
						entr.Concept = concFunction
						if isobject(GetSourceForAppFunction(el, entr)) then
							Set temp = GetSourceForAppFunction(el, entr)
						end if
                    End If
                    If (el.Stereotype = "ArchiMate_ApplicationComponent") Then
                        entr.Concept = concSystem
                        if isobject(GetSourceForAppComponent(el, entr)) then
							Set temp = GetSourceForAppComponent(el, entr)
						end if
                    End If
            End If
            If connector.SupplierID = gap.ElementID Then
                     Set el = repository.GetElementByID(connector.ClientID)
                     If (el.Stereotype = "ArchiMate_ApplicationService") Then
                        entr.AppService = el.Name
						entr.Concept = concService
                        Dim tempS As EA.Element
						if isobject(GetSourceForAppService(el, entr)) then
							Set tempS = GetSourceForAppService(el, entr)
						end if
                    End If
                    If (el.Stereotype = "ArchiMate_ApplicationFunction") Then
                        entr.AppFunction = el.Name
						entr.Concept = concFunction
                        Dim tempF As EA.Element
						if isobject(GetSourceForAppFunction(el, entr)) then
							Set tempF = GetSourceForAppFunction(el, entr)
						end if 
                    End If
                    If (el.Stereotype = "ArchiMate_ApplicationComponent") Then
                        Dim tempC As EA.Element
                        entr.Concept = concSystem
						if isobject(GetSourceForAppComponent(el, entr)) then
							Set tempC = GetSourceForAppComponent(el, entr)
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

Public Function GetSourceForAppService(ser, ByRef entr ) 
  
    Dim sourceFunction As EA.Element
    Dim sourceComponent As EA.Element
    Dim cn 
    ' call Session.Output ("verification of service" + ser.Name)
    For Each cn In ser.Connectors
        Dim el As EA.Element
        If cn.Stereotype = "ArchiMate_Composition" Then
            If cn.SupplierID = ser.ElementID Then
                  Set el = repository.GetElementByID(cn.ClientID)
				  ' call Session.Output ("app service" + ser.Name)
            End If
            If cn.ClientID = ser.ElementID Then
					Set el = repository.GetElementByID(cn.SupplierID)
                    ' call Session.Output ("getting source for app service " + ser.Name)
					if isobject(GetSourceForAppService(el, entr)) then
						Set sourceComponent = GetSourceForAppService(el, entr)
					end if
            End If
        End If
        If cn.Stereotype <> "ArchiMate_Realization" Then
        Else
             If cn.SupplierID = ser.ElementID Then
                    Set el = repository.GetElementByID(cn.ClientID)
                    If (el.Stereotype = "ArchiMate_ApplicationFunction") Then
					   ' call Session.Output ("getting 2 source for app service" + ser.Name)
                       if isobject(GetSourceForAppFunction(el, entr)) then
							Set sourceComponent = GetSourceForAppFunction(el, entr)
					   end if
                    End If
              End If
        End If
    Next 
    entr.AppService = ser.Name
	If isobject(sourceComponent)Then
        Set GetSourceForAppService = sourceComponent
    End If
	
End Function
Public Function GetSourceForAppFunction(fnc, ByRef entr) 

    Dim s As EA.Element
    Dim cnn
    For Each cnn In fnc.Connectors
        If cnn.Stereotype = "ArchiMate_Composition" Then
            Dim el As EA.Element
            If cnn.ClientID = fnc.ElementID Then
                    Set el = repository.GetElementByID(cnn.SupplierID)
				    ' call Session.Output ("source for app fuction " + fnc.Name)
					if isobject(GetSourceForAppFunction(el, entr)) then
					Set s = GetSourceForAppFunction(el, entr)
					end if
            End If
        End If
        If cnn.Stereotype = "ArchiMate_Assignment" Then
             Dim elX As EA.Element
             If cnn.SupplierID = fnc.ElementID Then
                 Set elX = repository.GetElementByID(cnn.ClientID)
                 ' call Session.Output ("source for app fuction " + fnc.Name)
                 if isobject(GetSourceForAppComponent(elX, entr, fnc)) then
					Set s = GetSourceForAppComponent(elX, entr, fnc)
				end if
            End If
             If cnn.ClientID = fnc.ElementID Then
                 Set elX = repository.GetElementByID(cnn.SupplierID)
                ' call Session.Output ("source for app fuction " + fnc.Name)
				' call Session.Output ("checking  " + elX.Name)
                 if isobject(GetSourceForAppComponent(elX, entr, fnc)) then
					Set s = GetSourceForAppComponent(elX, entr, fnc)
				 end if
            End If
            
        End If
    Next 
    entr.AppFunction = fnc.Name
    If not isobject(s)Then
        entr.AppComponent = "!!!!No parent component"
    Else
        Set GetSourceForAppFunction = s
    End If
End Function
Public Function GetSourceForAppComponent(comp, ByRef entr, ByRef fnc) 
	
    Dim cnnn 
    Dim sourceItself
    sourceItself = false
	Dim sourceAlreadyInMap
    sourceAlreadyInMap = false
	Dim tt
	'call Session.Output ("evaluating component - " + comp.Name)	
    if comp.Connectors.Count > 1 Then
		For Each cnnn In comp.Connectors
		sourceAlreadyInMap = false
		If cnnn.Stereotype = "ArchiMate_Composition" Then
			If cnnn.SupplierID = comp.ElementID Then
				sourceItself = True
				'call Session.Output ("Source for fnc - " + fnc.Name + " in comp " + comp.Name)
				Set s = comp
				Set GetSourceForAppComponent = s
				if componentForGap.Count > 0 then
					ss = componentForGap.Items
					for tt = 0 to componentForGap.Count - 1 
						'call Session.Output ("Verifying - " + comp.Name + "; entry " +CStr(entr.Id))
						if ss(tt).ElementId = comp.ElementId then
							'call Session.Output ("Component already in map - " + comp.Name + "; entry "+CStr(entr.Id))
							sourceAlreadyInMap = true
						end if
					Next
					if not sourceAlreadyInMap then 
						'call Session.Output ("Adding component - " + comp.Name + "; entry "+CStr(entr.Id))
						if not componentForGap.Exists(entr.Id) then
							componentForGap.Add entr.Id, comp
						end if
					End if
				else 
					'call Session.Output ("Adding component - " + comp.Name + "; entry "+CStr(entr.Id))
					if not componentForGap.Exists(entr.Id) then
						componentForGap.Add entr.Id, comp
					end if
				end if
			End If
			If cnnn.ClientID = comp.ElementID Then
				Dim el As EA.Element
				Set el = repository.GetElementByID(cnnn.SupplierID)
				sourceItself = False
				'call Session.Output ("Checking parent - " + el.Name)
				Set GetSourceForAppComponent = GetSourceForAppComponent(el, entr, fnc)
			end if
		Else
			if cnnn.Stereotype = "ArchiMate_Assignment" then
				If cnnn.SupplierID = comp.ElementID and cnnn.ClientID = fnc.ElementID Then
					'call Session.Output ("Checking comp " + comp.Name + " assignment for fnc - " + fnc.Name)
					if componentForGap.Count > 0 then
					ss = componentForGap.Items
					for tt = 0 to componentForGap.Count - 1 
						if ss(tt).ElementId = comp.ElementId then
							'call Session.Output ("Already in map- " + comp.Name)
							sourceAlreadyInMap = true
							Set GetSourceForAppComponent = comp
						end if
					Next
					if not sourceAlreadyInMap and not CompHasCompositionRel(comp) then 
						'call Session.Output ("Adding component - " + comp.Name)
						if not componentForGap.Exists(entr.Id) then
							componentForGap.Add entr.Id, comp
							Set s = comp
							Set GetSourceForAppComponent = comp
						end if
					End if
				else 				
					if not CompHasCompositionRel(comp) then
						'call Session.Output ("Adding component - " + comp.Name)
						if not componentForGap.Exists(entr.Id) then
							componentForGap.Add entr.Id, comp
							Set s = comp
							Set GetSourceForAppComponent = comp
						end if
					end if
				end if 
			end if
			sourceItself = True
		end if	
		End If
	Next
   Else
		if comp.Connectors.Count = 1 Then
			if comp.Connectors.GetAt(0).Stereotype = "ArchiMate_Composition" Then
				If comp.Connectors.GetAt(0).SupplierID = comp.ElementID Then
					sourceItself = True
					Set s = comp
					Set GetSourceForAppComponent = comp
					exit function
				End If
			End if
		end if
   End if 
 
  Dim keys, j, items
  keys = componentForGap.Keys
  items = componentForGap.Items
	for j = 0 to componentForGap.Count - 1 
		if keys(j) = entr.Id then
			Set GetSourceForAppComponent = componentForGap.Item(entr.Id)					
		end if
	Next
   If not isobject(GetSourceForAppComponent) Then
       entr.AppComponent = "!!!!No application component"
   Else
		if GetSourceForAppComponent.Name <> "" then
        entr.AppComponent = GetSourceForAppComponent.Name 
		else entr.AppComponent = "!!!!No application component"
		end if
    End If

End Function

Public function CompHasCompositionRel(comp) 
	Dim assignmentCompConnector
	if comp.Connectors.Count > 1 Then
		For Each assignmentCompConnector In comp.Connectors
			If assignmentCompConnector.Stereotype = "ArchiMate_Composition" Then
				CompHasCompositionRel = true
			else 
				CompHasCompositionRel = false
			end if
		next
	end if
	
End function
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

