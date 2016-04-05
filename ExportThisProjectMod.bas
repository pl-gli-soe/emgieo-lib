Attribute VB_Name = "ExportThisProjectMod"
' working great!

Global Const REPO_PATH = "C:\WORKSPACE\macros\MGO_LIB_REPO\"

Private Sub export_this_project()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_StdModule Then
            txt = VBComp.Name & ".bas"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            txt = VBComp.Name & ".cls"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            txt = VBComp.Name & ".frm"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        End If
         
    Next VBComp
    
    MsgBox "ready!"

End Sub
