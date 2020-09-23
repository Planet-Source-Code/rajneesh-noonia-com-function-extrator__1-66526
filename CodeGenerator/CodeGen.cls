VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public OutPutFileName As String

Public Sub CodeGenerate(ByVal pFileName As String)

Dim pApplication As New TLI.TLIApplication
Dim pTypeLibINfo As TLI.TypeLibInfo

Set pTypeLibINfo = pApplication.TypeLibInfoFromFile(pFileName)
Call GetSignatures(pTypeLibINfo)
'Call ExploreTypeLib(pTypeLibINfo)
End Sub


Private Function GetMemberInfoString(ByVal pMemberInfo As TLI.MemberInfo) As String
    Dim pOut As String
    Dim pParamInfo As ParameterInfo
    
    pOut = vbCrLf
    
    Select Case pMemberInfo.InvokeKind
        Case InvokeKinds.INVOKE_CONST:
            Debug.Assert False
        Case InvokeKinds.INVOKE_EVENTFUNC:
            Debug.Assert False
        Case InvokeKinds.INVOKE_FUNC:
            pOut = pOut & vbCrLf
            pOut = pOut & "'Function/Sub :" & pMemberInfo.HelpString & vbCrLf
            If pMemberInfo.ReturnType.VarType = VT_VOID Then
                pOut = pOut & "Public Sub " & pMemberInfo.Name & "("
            Else
                pOut = pOut & "Public Function " & pMemberInfo.Name & "("
            End If
            
            For Each pParamInfo In pMemberInfo.Parameters
               If pParamInfo.Optional Then
                    pOut = pOut & " Optional "
               End If
               pOut = pOut & "Byval p" & pParamInfo.Name & GetObjectTypeFromParameter(pParamInfo)
               pOut = pOut & ","
            Next
            
            If pMemberInfo.Parameters.Count > 0 Then
                pOut = Left(pOut, Len(pOut) - 1)
            End If
            
            pOut = pOut & ")"
            
            If pMemberInfo.ReturnType.VarType <> VT_VOID Then
                pOut = pOut & GetObjectTypeFromMember(pMemberInfo)
            End If
            pOut = pOut & vbCrLf
            pOut = pOut & "'" & vbCrLf
            
            If pMemberInfo.ReturnType.VarType = VT_VOID Then
                pOut = pOut & "End Sub" & vbCrLf
            Else
                pOut = pOut & "End Function" & vbCrLf
            End If
            
        Case InvokeKinds.INVOKE_PROPERTYGET:
            pOut = pOut & vbCrLf
            pOut = pOut & "'Property :" & pMemberInfo.HelpString & vbCrLf
            pOut = pOut & "Public Property Get " & pMemberInfo.Name & "()"
            pOut = pOut & GetObjectTypeFromMember(pMemberInfo)
            pOut = pOut & vbCrLf
            pOut = pOut & "'" & vbCrLf
            pOut = pOut & "End Property" & vbCrLf
        Case InvokeKinds.INVOKE_PROPERTYPUT:
            pOut = pOut & vbCrLf
            pOut = pOut & "'Property :" & pMemberInfo.HelpString & vbCrLf
            pOut = pOut & "Public Property Let " & pMemberInfo.Name & "("
            pOut = pOut & "Byval p" & pMemberInfo.Name & GetObjectTypeFromMember(pMemberInfo)
            pOut = pOut & ")" & vbCrLf
            pOut = pOut & "'" & vbCrLf
            pOut = pOut & "End Property" & vbCrLf
        Case InvokeKinds.INVOKE_PROPERTYPUTREF:
            pOut = pOut & vbCrLf
            pOut = pOut & "'Property :" & pMemberInfo.HelpString & vbCrLf
            pOut = pOut & "Public Property Set " & pMemberInfo.Name & "("
            pOut = pOut & "ByRef p" & pMemberInfo.Name & GetObjectTypeFromMember(pMemberInfo)
            pOut = pOut & ")" & vbCrLf
            pOut = pOut & "'" & vbCrLf
            pOut = pOut & "End Property" & vbCrLf
        Case InvokeKinds.INVOKE_UNKNOWN:
            Debug.Assert False
    End Select
    GetMemberInfoString = pOut & vbCrLf
End Function

Private Function GetObjectTypeFromParameter(ByVal pParameterInfo As TLI.ParameterInfo) As String
    Select Case pParameterInfo.VarTypeInfo
                Case TliVarType.VT_EMPTY:
                     GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as " & pParameterInfo.VarTypeInfo.TypeInfo.Name
                Case TliVarType.VT_I4: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Long "
                Case TliVarType.VT_BSTR: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  String "
                Case TliVarType.VT_DISPATCH: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Object "
                Case TliVarType.VT_BOOL: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Boolean "
                Case TliVarType.VT_VARIANT: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Variant "
                Case TliVarType.VT_I1: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Byte "
                Case TliVarType.VT_I2: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Integer "
                Case TliVarType.VT_INT: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Integer "
                Case TliVarType.VT_DATE: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Date "
                Case TliVarType.VT_DECIMAL: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " as  Decimal "
                Case TliVarType.VT_VOID: GetObjectTypeFromParameter = GetObjectTypeFromParameter & " "
            End Select
End Function


Private Function GetObjectTypeFromMember(ByVal pMemberInfo As TLI.MemberInfo) As String
    Select Case pMemberInfo.ReturnType.VarType
                Case TliVarType.VT_EMPTY:
                     GetObjectTypeFromMember = GetObjectTypeFromMember & " as " & pMemberInfo.ReturnType.TypeInfo.Name
                Case TliVarType.VT_I4: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Long "
                Case TliVarType.VT_BSTR: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  String "
                Case TliVarType.VT_DISPATCH: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Object "
                Case TliVarType.VT_BOOL: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Boolean "
                Case TliVarType.VT_VARIANT: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Variant "
                Case TliVarType.VT_I1: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Byte "
                Case TliVarType.VT_I2: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Integer "
                Case TliVarType.VT_INT: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Integer "
                Case TliVarType.VT_DATE: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Date "
                Case TliVarType.VT_DECIMAL: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  Decimal "
                Case TliVarType.VT_VOID: GetObjectTypeFromMember = GetObjectTypeFromMember & " "
                Case TliVarType.VT_UNKNOWN: GetObjectTypeFromMember = GetObjectTypeFromMember & " as  IUnknown "
            End Select
End Function

Private Sub GetSignatures(ByVal pTypeLibINfo As TLI.TypeLibInfo)
    Dim pTypeInfo As TLI.TypeInfo
    Dim pFso As New FileSystemObject
    Dim pTxtStream As TextStream
    Dim pMemberInfo As MemberInfo
    Dim pInterfaceInfo As TLI.InterfaceInfo
    
    Dim pOut As String
    Set pTxtStream = pFso.OpenTextFile(OutPutFileName, ForWriting, True)
    
    For Each pTypeInfo In pTypeLibINfo.TypeInfos
        pOut = ""
        Select Case pTypeInfo.TypeKind
            Case TypeKinds.TKIND_COCLASS:   'Public Creatable
                pOut = pOut & "CoClass:------------------ " & pTypeInfo.Name & String(5, vbTab) & IIf(Len(pTypeInfo.HelpString) > 0, "'" & pTypeInfo.HelpString, "") & vbCrLf
                For Each pInterfaceInfo In pTypeInfo.Interfaces
                    For Each pMemberInfo In pInterfaceInfo.Members
                        If pMemberInfo.MemberId < 1610612736 Then
                            pOut = pOut & GetMemberInfoString(pMemberInfo)
                        End If
                    Next
                Next
                pOut = pOut & "End CoClass" & vbCrLf & vbCrLf
            Case TypeKinds.TKIND_DISPATCH:  'Collections
                pOut = pOut & "Class:------------------ " & pTypeInfo.Name & String(5, vbTab) & IIf(Len(pTypeInfo.HelpString) > 0, "'" & pTypeInfo.HelpString, "") & vbCrLf
                If Left(pTypeInfo.Name, 1) <> "_" Then
                    For Each pMemberInfo In pTypeInfo.Members
                        If pMemberInfo.MemberId < 1610612736 Then
                            pOut = pOut & GetMemberInfoString(pMemberInfo)
                        End If
                    Next
                End If
                pOut = pOut & "End Class" & vbCrLf & vbCrLf
            Case TypeKinds.TKIND_ENUM:
                pOut = pOut & "Public Enum " & pTypeInfo.Name & String(5, vbTab) & IIf(Len(pTypeInfo.HelpString) > 0, "'" & pTypeInfo.HelpString, "") & vbCrLf
                For Each pMemberInfo In pTypeInfo.Members
                    pOut = pOut & vbTab & pMemberInfo.Name & " = " & pMemberInfo.Value & vbCrLf
                Next
                pOut = pOut & "End Enum" & vbCrLf & vbCrLf
            Case TypeKinds.TKIND_INTERFACE:
                pOut = pOut & "Interface :------------------" & pTypeInfo.Name & String(5, vbTab) & IIf(Len(pTypeInfo.HelpString) > 0, "'" & pTypeInfo.HelpString, "") & vbCrLf
                For Each pInterfaceInfo In pTypeInfo.Interfaces
                    For Each pMemberInfo In pInterfaceInfo.Members
                        If pMemberInfo.MemberId < 1610612736 Then
                            pOut = pOut & GetMemberInfoString(pMemberInfo)
                        End If
                    Next
                Next
                pOut = pOut & "End Interface" & vbCrLf & vbCrLf
            Case Else
                Debug.Assert False
                Debug.Print "ELSE"
        End Select
        If Len(pOut) > 0 Then pTxtStream.Write pOut
    Next
    pTxtStream.Close
    Set pTxtStream = Nothing
    Set pFso = Nothing
End Sub



