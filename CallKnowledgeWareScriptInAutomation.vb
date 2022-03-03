Private Const sTemplateShape As String = "3D_Prd_00101074"
Private Const sOldInstName As String = "Physical Product00133699.1"
Private Const sNewRefName As String = "ZylinderBWH"

Sub CATMain()

Dim myRelations As Relations
Dim myLaw As Law
Dim pInput As Parameter

'Open Template 3DShape
GetOpenTemplateShape

'run the law by setting the String.1 Parameter. If value of String.1 is still the same, law will be activated - deactivated
Set pInput = CATIA.ActiveEditor.ActiveObject.Parameters.GetItem("String.1")
If Not pInput.Value = sOldInstName + "|" + sNewRefName Then
    pInput.Value = sOldInstName + "|" + sNewRefName
Else
    Set myLaw = CATIA.ActiveEditor.ActiveObject.Relations.Item(1)
    If myLaw.Activated = True Then myLaw.Deactivate
    If myLaw.Activated = False Then myLaw.Activate
End If

'save the 3DShape Template
Dim oPLMPropagateService As PLMPropagateService
Set oPLMPropagateService = CATIA.GetSessionService("PLMPropagateService")
oPLMPropagateService.PLMPropagate

CATIA.ActiveWindow.Close


End Sub

Private Sub GetOpenTemplateShape()

     Dim aSearch As SearchService
     Dim DBSearch As DatabaseSearch
     Dim SearchString As String
     Dim strSearch As String
     
     Dim oEntities As PLMEntities
     Dim OpenService As PLMOpenService
     Dim oEditor As Editor
     
     ' set search service and define search
     Set aSearch = CATIA.GetSessionService("Search")
     Set DBSearch = aSearch.DatabaseSearch
     DBSearch.Mode = SearchMode_Extended
     DBSearch.BaseType = "3DShape"
     DBSearch.AddExtendedCriteria "PLM_ExternalID", sTemplateShape, SearchOperator_EQ
    
    ' execute search
     aSearch.Search

    'Open search result
    Set oEntities = DBSearch.Results
    Set OpenService = CATIA.GetSessionService("PLMOpenService")
    OpenService.PLMOpen oEntities.Item(1), oEditor
   
End Sub

