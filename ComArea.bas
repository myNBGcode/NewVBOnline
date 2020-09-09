Attribute VB_Name = "ComArea"
Option Explicit

Public Function DeclareComArea_(name As String, comareaid As String, Method As String, TrnId As String, filename As String, InputName As String, OutputName As String, Optional StructContainer As Buffers) As cXmlComArea

    Dim ComArea As cXmlComArea
    Set ComArea = New cXmlComArea
    Set ComArea.content = CreateElementComArea(name, comareaid, Method, TrnId, filename, InputName, OutputName)
    If IsMissing(StructContainer) Then
        Set StructContainer = GenWorkForm.AppBuffers
    End If
    If StructContainer Is Nothing Then
        Set StructContainer = GenWorkForm.AppBuffers
    End If
    Set ComArea.Container = StructContainer
    If DatabaseMdl.BuildComArea(StructContainer, comareaid, filename) Then
     
    Else
       'error
       'comarea.owner.TrnBuffers.ByName(comareaid).ClearData
    End If

    Set DeclareComArea_ = ComArea
    
End Function

Private Function CreateElementComArea(name As String, comareaid As String, Method As String, TrnId As String, filename As String, InputName As String, OutputName As String) As IXMLDOMElement
    Dim doc As MSXML2.DOMDocument30
     Set doc = New MSXML2.DOMDocument30
    Dim elm As IXMLDOMElement
    Set elm = doc.createElement("root")
    
    doc.appendChild doc.createElement("root")
    doc.documentElement.appendChild doc.createElement("comarea")
    
    Dim area As IXMLDOMElement
    Set area = doc.documentElement.selectSingleNode("comarea")
    Dim attr As IXMLDOMAttribute
    Set attr = doc.createAttribute("name")
    attr.value = name
    area.setAttributeNode attr
    
    Set attr = doc.createAttribute("id")
    attr.value = comareaid
    area.setAttributeNode attr
    
    Set attr = doc.createAttribute("filename")
    attr.value = filename
    area.setAttributeNode attr
    
    area.appendChild doc.createElement("method")
    
    Set area = doc.selectSingleNode("//method")
    
    Set attr = doc.createAttribute("name")
    attr.value = Method
    area.setAttributeNode attr
    
    Set attr = doc.createAttribute("trncall")
    attr.value = TrnId
    area.setAttributeNode attr
    
    Set attr = doc.createAttribute("inputname")
    attr.value = InputName
    area.setAttributeNode attr
    
    Set attr = doc.createAttribute("outputname")
    attr.value = OutputName
    area.setAttributeNode attr
    
    Set CreateElementComArea = doc.documentElement.selectSingleNode("comarea")
End Function

