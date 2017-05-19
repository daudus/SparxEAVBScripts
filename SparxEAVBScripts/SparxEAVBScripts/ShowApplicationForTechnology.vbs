option explicit
!INC Local Scripts.EAConstants-VBScript
'
'ScriptName:
'Author:
'Purpose:
'Date:
'

sub main

	Dim element,c,s
	Dim package
	Dim entries
	Dim connector
	Dim selectedObjectType
	Dim cc,ss
	Dim i

	i = 0
	Set entries=CreateObject("Scripting.Dictionary")
	Set element=Repository.GetContextObject()
	selectedObjectType=Repository.GetContextItemType()
	if selectedObjectType <> otElement then
		call Session.Prompt("You have to select element",0)
		exit sub
	end if
	call Session.Prompt("Going through connectors",0)
	For Each connector In element.Connectors
		cc = connector.ClientID
		ss = connector.SupplierID
		call Session.Output("Connectorstereotype: "+connector.Stereotype)
		call Session.Output("ClientID: " + Cstr(cc))
		call Session.Output("SupplierID: " + Cstr(ss))
		set c = Repository.GetElementByID(cc)
		set s = Repository.GetElementByID(connector.SupplierID)
		call Session.Output("Client Name: " + c.Name)
		call Session.Output("Supplier Name: " + s.Name)
		i = i +1
	Next
	call Session.Prompt("FINISHED! " & i & " connectors found.",0)
end sub

main
