Imports System.Runtime.InteropServices
Module Module1

    Sub Main()
        Dim objApp As SolidEdgeFramework.Application
        Dim objPartDoc As SolidEdgePart.PartDocument
        Dim objProfileSets As SolidEdgePart.ProfileSets
        Dim objprofileSet As SolidEdgePart.ProfileSet
        Dim objProfiles As SolidEdgePart.Profiles
        Dim objProfile As SolidEdgePart.Profile
        Dim objTextProfiles As SolidEdgeFrameworkSupport.TextProfiles
        Dim objTextProfile As SolidEdgeFrameworkSupport.TextProfile
        Dim TextProfileFound As Boolean = False

        Try
            objApp = Marshal.GetActiveObject("SolidEdge.Application")
        Catch ex As Exception
            MsgBox("Unable to connect to Solid Edge!")
        End Try

        objPartDoc = objApp.ActiveDocument
        objProfileSets = objPartDoc.ProfileSets
        'loop through the profile sets
        For i As Integer = 1 To objProfileSets.Count
            objprofileSet = objProfileSets.Item(i)
            objProfiles = objprofileSet.Profiles
            'check if the profile set has a profile
            If objProfiles.Count > 0 Then
                objProfile = objProfiles.Item(1)
                objTextProfiles = objProfile.TextProfiles
                'check if the profile has a text profile
                If objTextProfiles.Count = 1 Then
                    objTextProfile = objTextProfiles.Item(1)
                    objTextProfile.Text = "%{Document Number}/%{Revision number}"
                    TextProfileFound = True
                    Exit For
                End If
            End If
        Next
        'notify user if no text profile was found
        If TextProfileFound = False Then
            MsgBox("No text profile found! Ensure part has a text profile then run program again.")
            Exit Sub
        End If

        MsgBox("Text profile successfully updated!")

    End Sub

End Module
