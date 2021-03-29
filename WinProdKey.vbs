Option Explicit

Dim service
Dim ser, obj
Dim serout, strOut
Dim lic, licStat
Dim objshell, path, ProductData
Dim serFind, objFind

serFind = "Version, RemainingWindowsReArmCount"
objFind = "ProductKeyChannel, ID, ApplicationId, ProductKeyID, OfflineInstallationId, LicenseStatus, GracePeriodRemaining, PartialProductKey, LicenseIsAddon, Description"

Set service = GetObject("winmgmts:\\.\root\cimv2")
Set objshell = CreateObject("WScript.Shell")

For Each ser in service.ExecQuery("SELECT " & serFind & " FROM SoftwareLicensingService")
	serOut = "Software licensing service version: " & ser.Version & vbNewLine & "Remaining Winows rearm count: " & ser.RemainingWindowsReArmCount
Next

For Each obj in service.ExecQuery("SELECT " & objFind & " FROM SoftwareLicensingProduct")
	if (GetIsPrimaryWindowsSKU(obj)) Then
		lic = obj.LicenseStatus
		Select case lic
			case 1
			licStat = "Licensed"
			case 2
			licStat = "Initial grace period"
			case 3
			licStat = "Additional grace period (KMS license expired or ardware out of tolerance)"
			case 4
			licStat = "Non-genuine grace period"
			case 5
			licStat = "Notification"
			case 6
			licStat = "Extended grace period"
			case else
			licStat = "Unknown"
		End Select
		strOut = "Product Channel: " & obj.ProductKeyChannel & vbNewLine & vbNewLine & "Activation ID: " & obj.ID & vbNewLine & "Application ID: " & obj.ApplicationId & vbNewLine & "Extended PID: " & obj.ProductKeyID & vbNewLine & "Offline Installation ID: " & obj.OfflineInstallationId & vbNewLine & vbNewLine & "License Status: " & licStat & vbNewLine & "Grace period remaining: " & obj.GracePeriodRemaining & vbNewLine 
	End If
Next

'Set registry key path
Path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
'Output
ProductData = "Product Name: " & objshell.RegRead(Path & "ProductName") & vbNewLine & "Product ID: " & objshell.RegRead(Path & "ProductID") & vbNewLine & "Installed Key: " & ConvertToKey(objshell.RegRead(Path & "DigitalProductId")) & vbNewLine & strOut & vbNewLine & serOut & vbNewLine & "Time: " & Now & vbnewline
'Show messbox if save to a file
Save(ProductData)
' If vbYes = MsgBox(ProductData & vblf & vblf & "Save to a file?", vbYesNo + vbQuestion, "BackUp Windows Key Information") then
	' Save(ProductData)
' End If
'Convert binary to chars
Function ConvertToKey(Key)
	Const KeyOffset = 52
	Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert
'Check if OS is Windows 8
	isWin8 = (Key(66) \ 6) And 1
	Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
	i = 24
	Maps = "BCDFGHJKMPQRTVWXY2346789"
	Do
		Current= 0
		j = 14
		Do
			Current = Current* 256
			Current = Key(j + KeyOffset) + Current
			Key(j + KeyOffset) = (Current \ 24)
			Current=Current Mod 24
			j = j -1
		Loop While j >= 0
		i = i -1
		KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput
		Last = Current
	Loop While i >= 0
	
	If (isWin8 = 1) Then
		keypart1 = Mid(KeyOutput, 2, Last)
		insert = "N"
		KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
		If Last = 0 Then KeyOutput = insert & KeyOutput
	End If
	ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)
End Function

'Save data to a file
Function Save(Data)
	Dim fso, txt, clock
	clock = Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "@" & Hour(Now) & ";" & Minute(Now)
'Create a text file in location of script
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set txt = fso.CreateTextFile(fso.GetAbsolutePathName("") & "\WinKey_" & clock & ".txt")
	txt.Writeline Data
	txt.Close
End Function

Function GetIsPrimaryWindowsSKU(objProduct)
    Dim iPrimarySku
    Dim bIsAddOn

    'Assume this is not the primary SKU
    iPrimarySku = 0
    If (LCase(objProduct.ApplicationId) = "55c92734-d682-4d71-983e-d6ec3f16059f" And objProduct.PartialProductKey <> "") Then
        'If we can get verify the AddOn property then we can be certain
        On Error Resume Next
        bIsAddOn = objProduct.LicenseIsAddon
        If Err.Number = 0 Then
            If bIsAddOn = true Then
                iPrimarySku = 0
            Else
                iPrimarySku = 1
            End If
        Else
            'If we can not get the AddOn property then we assume this is a previous version
            'and we return a value of Uncertain, unless we can prove otherwise
            If (IsKmsClient(objProduct.Description) Or IsKmsServer(objProduct.Description)) Then
                'If the description is KMS related, we can be certain that this is a primary SKU
                iPrimarySku = 1
            Else
                'Indeterminate since the property was missing and we can't verify KMS
                iPrimarySku = 2
            End If
        End If
    End If
    GetIsPrimaryWindowsSKU = iPrimarySku
End Function