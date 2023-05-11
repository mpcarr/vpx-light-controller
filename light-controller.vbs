'***********************************************************************************************************************
' Lights State Controller - 0.5.1
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************

Dim lightCtrl : Set lightCtrl = new LStateController

Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqOverride, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval

    Private Sub Class_Initialize()
        Set m_lights = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_off = CreateObject("Scripting.Dictionary")
        Set m_seqRunners = CreateObject("Scripting.Dictionary")
        Set m_currentFrameState = CreateObject("Scripting.Dictionary")
        Set m_seqs = CreateObject("Scripting.Dictionary")
        Set m_pulse = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_seqOverride = new LCSeqRunner
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
        m_seqOverride.name = "lSeqLightsOverride"
		m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
    End Sub

    Private Sub m_assignStateForFrame(key, state)
        If m_currentFrameState.Exists(key) Then
            m_currentFrameState.Remove key
        End If
        m_currentFrameState.Add key, state
    End Sub

    Public Sub LoadLightShows()
        Dim oFile
        Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-out.txt",2,true)
        For Each oFile In oFSO.GetFolder(cGameName & "_LightShows").Files
            If LCase(oFSO.GetExtensionName(oFile.Name)) = "yaml" And Not Left(oFile.Name,6) = "lights" Then
                Dim textStream : Set textStream = oFSO.OpenTextFile(oFile.Path, 1)
                Dim show : show = textStream.ReadAll
                Dim fileName : fileName = "lSeq" & Replace(oFSO.GetFileName(oFile.Name), "."&oFSO.GetExtensionName(oFile.Name), "")
                Dim lcSeq : lcSeq = "Dim " & fileName & " : Set " & fileName & " = New LCSeq"&vbCrLf
                lcSeq = lcSeq + fileName & ".Name = """&fileName&""""&vbCrLf
                Dim seq : seq = ""
                Dim re : Set re = New RegExp
                With re
                    .Pattern    = "- time:.*?\n"
                    .IgnoreCase = False
                    .Global     = True
                End With
                Dim matches : Set matches = re.execute(show)
                Dim steps : steps = matches.Count
                Dim match, nextMatchIndex, uniqueLights
                Set uniqueLights = CreateObject("Scripting.Dictionary")
                nextMatchIndex = 1
                For Each match in matches
                    Dim lightStep
                    If Not nextMatchIndex < steps Then
                        lightStep = Mid(show, match.FirstIndex, Len(show))
                    Else
                        lightStep = Mid(show, match.FirstIndex, matches(nextMatchIndex).FirstIndex - match.FirstIndex)
                        nextMatchIndex = nextMatchIndex + 1
                    End If

                    Dim re1 : Set re1 = New RegExp
                    With re1
                        .Pattern        = ".*:?: '([A-Fa-f0-9]{6})'"
                        .IgnoreCase     = True
                        .Global         = True
                    End With

                    Dim lightMatches : Set lightMatches = re1.execute(lightStep)
                    If lightMatches.Count > 0 Then
                        Dim lightMatch, lightStr
                        lightStr = "Array("
                        For Each lightMatch in lightMatches
                            Dim sParts : sParts = Split(lightMatch.Value, ":")
                            Dim lightName : lightName = Trim(sParts(0))
                            Dim color : color = Trim(Replace(sParts(1),"'", ""))
                            If color = "000000" Then
                                lightStr = lightStr + """"&lightName&"|0|000000"","
                            Else
                                lightStr = lightStr + """"&lightName&"|100|"&color&""","
                            End If
                            uniqueLights(lightname) = 0
                        Next
                        lightStr = Left(lightStr, Len(lightStr) - 1)
                        lightStr = lightStr & ")"
                        seq = seq + lightStr & ", _"&vbCrLf
                    Else
                        seq = seq + "Array(), _"&vbCrLf
                    End If

                    
                    Set re1 = Nothing
                Next
                
                lcSeq = lcSeq + filename & ".Sequence = Array( " & Left(seq, Len(seq) - 5) & ")"&vbCrLf
                'lcSeq = lcSeq + seq & vbCrLf
                lcSeq = lcSeq + fileName & ".UpdateInterval = 20"&vbCrLf
                lcSeq = lcSeq + fileName & ".Color = Null"&vbCrLf
                lcSeq = lcSeq + fileName & ".Repeat = False"&vbCrLf

                'MsgBox(lcSeq)
                objFileToWrite.WriteLine(lcSeq)
                ExecuteGlobal lcSeq
                Set re = Nothing

                textStream.Close
            End if
        Next
        'Clean up
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Set oFile = Nothing
        Set oFSO = Nothing
    End Sub

    Public Sub CompileLights(collection, name)
        Dim light
        Dim lights : lights = "light:" & vbCrLf
        For Each light in collection
            lights = lights + light.name & ":"&vbCrLf
            lights = lights + "   x: "& light.x/tablewidth & vbCrLf
            lights = lights + "   y: "& light.y/tableheight & vbCrLf
        Next
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-"&name&".yaml",2,true)
	    objFileToWrite.WriteLine(lights)
	    objFileToWrite.Close
	    Set objFileToWrite = Nothing
        Debug.print("Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml")
    End Sub

    Public Sub RegisterLights()
        Dim lampzIdx
        For lampzIdx = 0 to UBound(Lampz.obj)
            If Lampz.IsLight(lampzIdx) Then
                Dim lcItem : Set lcItem = new LCItem
                Dim vpxLight
                If IsArray(Lampz.obj(lampzIdx)) Then
                    Dim tmp : tmp = Lampz.obj(lampzIdx)
                    Set vpxLight = tmp(0)
                Else
                    Set vpxLight = Lampz.obj(lampzIdx)
                    
                End If
                Lampz.Modulate(lampzIdx) = 1/100
                Lampz.FadeSpeedUp(lampzIdx) = 100/30 : Lampz.FadeSpeedDown(lampzIdx) = 100/120
                lcItem.Init lampzIdx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                
                m_lights.Add vpxLight.Name, lcItem
                m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
            End If
        Next        
    End Sub

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        m_LightOn(light.name)
    End Sub

    Public Sub LightOnWithColor(light, color)
        m_LightOnWithColor light.name, color
    End Sub

    Public Sub FlickerOn(light)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            m_lightOn(name)

            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, Array(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1)
        End If
    End Sub  
    
    Public Sub LightColor(light, color)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Color = color
            End If

        End If
    End Sub

    Private Sub m_LightOn(name)
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Private Sub m_LightOnWithColor(name, color)
        If m_lights.Exists(name) Then
            m_lights(name).Color = color
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If

            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Public Sub LightOff(light)
        m_lightOff(light.name)
    End Sub

    Private Sub m_lightOff(name)
        If m_lights.Exists(name) Then
            If m_on.Exists(name) Then 
                m_on.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_off.Exists(name) Then 
                Exit Sub
            End If
            m_off.Add name, m_lights(name)
        End If
    End Sub

    Public Sub Pulse(light, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, Array(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount)
        End If
    End Sub

    Public Sub PulseWithProfile(light, profile, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, Array(m_lights(name), profile, 0, m_pulseInterval, repeatCount)
        End If
    End Sub       

    Public Sub LightLevel(light, lvl)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Level = lvl

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light)
            End If
        End If
    End Sub

    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name) Then
                m_seqs(name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add name, seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub RemoveShot(name, light)
        If m_lights.Exists(light.name) And m_seqs.Exists(name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveItem m_seqs(name)
            If IsNUll(m_seqRunners("lSeqRunner"&CStr(light.name)).CurrentItem) Then
               LightOff(light)
            End If
        End If
    End Sub

    Public Sub RemoveAllShots()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            m_assignStateForFrame light, Array(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub RemoveShotsFromLight(light)
        If m_lights.Exists(light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveAll   
            m_lightOff(light.name)  
        End If
    End Sub

    Public Sub Blink(light)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Function GetLightState(light)
        GetLightState = 0
        If(m_lights.Exists(light.name)) Then
            If m_on.Exists(light.name) Then
                GetLightState = 1
            Else
                If m_seqs.Exists(light.name & "Blink") Then
                    GetLightState = 2
                End If
            End If
        End If
    End Function

    Public Function IsShotLit(name, light)
        IsShotLit = False
        If(m_lights.Exists(light.name)) Then
            If m_seqRunners("lSeqRunner"&CStr(light.name)).HasSeq(name) Then
                IsShotLit = True
            End If
        End If
    End Function

    Public Sub CreateSeqRunner(name)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).AddItem lcSeq
    End Sub

    Public Sub RemoveLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).RemoveItem lcSeq
    End Sub

    Public Sub RemoveAllLightSeq(lcSeqRunner)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).RemoveAll
    End Sub

    Public Sub AddTableLightSeq(lcSeq)
        If IsNull(m_seqOverride.CurrentItem) Then
            Dim light
            For Each light in m_lights.Keys()
                m_assignStateForFrame light, Array(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverride.AddItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(lcSeq)
        m_seqOverride.RemoveItem lcSeq
        If IsNull(m_seqOverride.CurrentItem) Then
            Dim light
            For Each light in m_lights.Keys()
                m_assignStateForFrame light, Array(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        m_seqOverride.RemoveAll
        Dim light
		For Each light in m_lights.Keys()
            m_assignStateForFrame light, Array(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub SyncWithVpxLights(lightSeq)
        Execute "m_vpxLightSyncCollection = ColToArray(" & CStr(lightSeq.Collection) & ")"
        m_vpxLightSyncRunning = True
    End Sub

    Public Sub StopSyncWithVpxLights()
        m_vpxLightSyncRunning = False
        m_vpxLightSyncClear = True
		m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
    End Sub

	Public Sub SetVpxSyncLightColor(color)
		m_tableSeqColor = color
	End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
		m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
	End Sub

    Public Sub UseToolkitColoredLightMaps()
        
        Dim sUpdateLightMap
        sUpdateLightMap = "Sub UpdateLightMap(idx, lightmap, intensity, ByVal aLvl)" + vbCrLf    
        sUpdateLightMap = sUpdateLightMap + "   if Lampz.UseFunc then aLvl = Lampz.FilterOut(aLvl)	'Callbacks don't get this filter automatically" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   lightmap.Opacity = aLvl * intensity" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   If IsArray(Lampz.obj(idx) ) Then" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.Color = Lampz.obj(idx)(0).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   Else" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.color = Lampz.obj(idx).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   End If" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "End Sub" + vbCrLf

        ExecuteGlobal sUpdateLightMap

        Dim x
        For x=0 to Ubound(Lampz.cCallback)
            Lampz.cCallback(x) = Replace(Lampz.cCallback(x), "UpdateLightMap ", "UpdateLightMap " & x & ",")
            Lampz.Callback(x) = "" 'Force Callback Sub to be build
        Next
    End Sub

    Private Function m_buildBlinkSeq(light)
        Dim i, buff : buff = Array()
        ReDim buff(Len(light.BlinkPattern)-1)
        For i = 0 To Len(light.BlinkPattern)-1
            
            If Mid(light.BlinkPattern, i+1, 1) = 1 Then
                buff(i) = light.name & "|100"
            Else
                buff(i) = light.name & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Public Sub Update()

		m_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
		Dim x
        Dim lk
        dim color
        Dim lightKey
        Dim lcItem
        Dim tmpLight

        If Not IsNull(m_seqOverride.CurrentItem) Then
            RunLightSeq m_seqOverride, "lightsOverride"
        Else
            If HasKeys(m_on) Then   
                For Each lightKey in m_on.Keys()
                    Set lcItem = m_on(lightKey)
                    m_assignStateForFrame lightKey, Array(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
                    m_assignStateForFrame lightKey, Array(m_pulse(lightKey)(1)(m_pulse(lightKey)(2)), m_pulse(lightKey)(0).Color, m_pulse(lightKey)(0).Idx)
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey)(3) - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey)(2)
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If
                    
                    Dim pulseCount : pulseCount = m_pulse(lightKey)(4)
                    
                    If pulseIdx > UBound(m_pulse(lightKey)(1)) Then
                        If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse(lightKey) = Array(m_pulse(lightKey)(0),m_pulse(lightKey)(1), pulseIdx, pulseUpdateInt, pulseCount)
                        Else
                            m_pulse(lightKey) = Array(m_pulse(lightKey)(0),m_pulse(lightKey)(1), pulseIdx, pulseUpdateInt, 0)
                            m_pulse.Remove lightKey
                        End If
                    Else
                        m_pulse(lightKey) = Array(m_pulse(lightKey)(0),m_pulse(lightKey)(1), pulseIdx, pulseUpdateInt, pulseCount)
                    End If
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    m_assignStateForFrame lightKey, Array(0, Null, lcItem.Idx)
                Next
            End If

            If HasKeys(m_seqRunners) Then
                Dim k
                For Each k in m_seqRunners.Keys()
                    Dim lsRunner: Set lsRunner = m_seqRunners(k)
                    If Not IsNull(lsRunner.CurrentItem) Then
                            RunLightSeq lsRunner, k
                    End If
                Next
            End If

            If m_vpxLightSyncRunning = True Then
                Dim lx
                For Each lx in m_vpxLightSyncCollection
                    'sync each light being ran by the vpx LS
                    dim syncLight : syncLight = Null
                    If m_lights.Exists(lx.name) Then
                        'found a light
                        Set syncLight = m_lights(lx.name)
                    End If
                    If Not IsNull(syncLight) Then
                        'Found a light to sync.
                        Dim lightState

                        If IsNull(m_tableSeqColor) Then
                            color = syncLight.Color
                        Else
                            If Not IsArray(m_tableSeqColor) Then
                                color = Array(m_TableSeqColor, Null)
                            Else
                                color = m_tableSeqColor
                            End If
                        End If

                        If Not IsNull(m_tableSeqFadeUp) Then
                            Lampz.FadeSpeedUp(syncLight.Idx) = m_tableSeqFadeUp
                        End If
                        If Not IsNull(m_tableSeqFadeDown) Then
                            Lampz.FadeSpeedDown(syncLight.Idx) = m_tableSeqFadeDown
                        End If
                        
                        
                        If IsArray(Lampz.obj(syncLight.Idx)) Then 
                            Set tmpLight = Lampz.obj(syncLight.Idx)(0)
                        Else
                            Set tmpLight = Lampz.obj(syncLight.Idx)
                        End If
                        m_assignStateForFrame syncLight.name, Array(tmpLight.GetInPlayState*100,color, syncLight.Idx)                     
                    End If
                Next
            End If

            If m_vpxLightSyncClear = True Then  
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lk in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncClearLight : syncClearLight = Null
                        If m_lights.Exists(lk.name) Then
                            'found a light
                            Set syncClearLight = m_lights(lk.name)
                        End If
                        If Not IsNull(syncClearLight) Then
                            m_assignStateForFrame syncClearLight.name, Array(0, syncClearLight.Color, syncClearLight.idx) 
                            'Lampz.state(syncClearLight.idx) = 0
                            Lampz.FadeSpeedUp(syncClearLight.Idx) = 100/30
                            Lampz.FadeSpeedDown(syncClearLight.Idx) = 100/120
                        End If 
                    Next
                End If
               
                m_vpxLightSyncClear = False
            End If
        End If
        

        If HasKeys(m_currentFrameState) Then
			
            Dim frameStateKey
            For Each frameStateKey in m_currentFrameState.Keys()
                Dim idx : idx = m_currentFrameState(frameStateKey)(2)
                'Debug.print("Changing light idx: " & CStr(m_currentFrameState(frameStateKey)(2)) & " -> " & CStr(m_currentFrameState(frameStateKey)(0)) & ". FrameTime: " & m_frametime)
                
                Dim newColor : newColor = m_currentFrameState(frameStateKey)(1)

                If Not IsNull(newColor) Then
                    'Debug.Print("Updating color")
                    'Check current lampz color is the new color coming in, if not, set the new color.
                    
				    If IsArray(Lampz.obj(idx) ) Then	'if array
                        Set tmpLight = Lampz.obj(idx)(0)
                    Else
                        Set tmpLight = Lampz.obj(idx)
                    End If
	
					Dim c, cf, bUpdate
					c = newColor(0)
					cf= newColor(1)

					If Not IsNull(c) Then
						If Not CStr(tmpLight.Color) = CStr(c) Then
							bUpdate = True
						End If
					End If

					If Not IsNull(cf) Then
						If Not CStr(tmpLight.ColorFull) = CStr(cf) Then
							bUpdate = True
						End If
					End If


                    If bUpdate Then
                        'Update lampz color
                        If IsArray(Lampz.obj(idx)) Then

                            for each x in Lampz.obj(idx)
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                            Next
                        Else
                            If Not IsNull(c) Then
                                Lampz.obj(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lampz.obj(idx).colorFull = cf
                            End If
                        End If
                    End If
                        
                    
                    'If Lampz.State(idx) = m_currentFrameState(frameStateKey)(0) Then
                        'Debug.print("Forcing callbacks")
                        If Lampz.UseCallBack(idx) then Proc Lampz.name & idx,Lampz.Lvl(idx)*Lampz.Modulate(idx)	'Proc
                    'End If'force object updates (callbacks)
            	End If
                
                Lampz.state(idx) = CInt(m_currentFrameState(frameStateKey)(0)) 'Lampz will handle redundant updates
				 
            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HasKeys(o)
        Dim Success
        Success = False

        On Error Resume Next
            o.Keys()
            Success = (Err.Number = 0)
        On Error Goto 0
        HasKeys = Success
    End Function

    Private Sub RunLightSeq(seqRunner, k)

        Dim lcSeq: Set lcSeq = seqRunner.CurrentItem
        dim lsName
        
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        Else
            Dim framesRemaining, seq, color
            seq = lcSeq.Sequence


            Dim name
            Dim ls, x
            If IsArray(seq(lcSeq.CurrentIdx)) Then
                For x = 0 To UBound(seq(lcSeq.CurrentIdx))
                    lsName = Split(seq(lcSeq.CurrentIdx)(x),"|")
                    name = lsName(0)
                    If m_lights.Exists(name) Then
                        Set ls = m_lights(name)
                        
						color = lcSeq.Color

                        If IsNull(color) Then
							'Debug.Print("seq color null")
							color = ls.Color
                        End If
						
                        If Ubound(lsName) = 2 Then
							'Debug.Print(lsName(0) & ":" &lsName(2))
                            m_assignStateForFrame name, Array(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        Else
                            m_assignStateForFrame name, Array(lsName(1), color, ls.Idx)
                        End If
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
					color = lcSeq.Color
                    If IsNull(color) Then
                        color = ls.Color
                    End If
                    If Ubound(lsName) = 2 Then
                        m_assignStateForFrame name, Array(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), Null), ls.Idx)
                    Else
                        m_assignStateForFrame name, Array(lsName(1), color, ls.Idx)
                    End If
                End If
            End If
            framesRemaining = lcSeq.Update(m_frameTime)
            'Debug.print(framesRemaining)
            If framesRemaining < 0 Then
                'Debug.print("Advancing")
                lcSeq.ResetInterval()
                lcSeq.NextFrame()
            End If
            
        End If
    End Sub

End Class


Class LCItem
	
	Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get Color()
            Color=m_color
        End Property

        Public Property Let Color(input)
            If IsNull(input) Then
				m_Color = Null
			Else
				If Not IsArray(input) Then
					input = Array(input, null)
				End If
				m_Color = input
			End If
	    End Property

        Public Property Let Level(input)
            m_level = input
	    End Property

        Public Property Get Level()
            Level=m_level
        End Property

        Public Property Get Name()
            Name=m_name
        End Property

        Public Property Get X()
            X=m_x
        End Property

        Public Property Get Y()
            Y=m_y
        End Property

        Public Sub Init(idx, intervalMs, color, name, x, y)
            m_Idx = idx
            If Not IsArray(color) Then
                m_color = Array(color, null)
            Else
                m_color = color
            End If
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
	    End Sub

End Class

Class LCSeq
	
	Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
        m_currentIdx = input
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
	Public Property Let Sequence(input)
		m_sequence = input
	End Property

    Public Property Get Color()
        Color=m_color
    End Property
    
	Public Property Let Color(input)
		If IsNull(input) Then
			m_Color = Null
		Else
			If Not IsArray(input) Then
				input = Array(input, null)
			End If
			m_Color = input
		End If
	End Property

    Public Property Get Name()
        Name=m_name
    End Property
    
	Public Property Let Name(input)
		m_name = input
	End Property        

    Public Property Get UpdateInterval()
        UpdateInterval=m_updateInterval
    End Property

    Public Property Let UpdateInterval(input)
        m_updateInterval = input
        ResetInterval()
    End Property

    Public Property Get Repeat()
        Repeat=m_repeat
    End Property

    Public Property Let Repeat(input)
        m_repeat = input
    End Property


    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Array(Null, Null)
        m_updateInterval = 180
        m_repeat = False
        m_Frames = 180
    End Sub

    Public Property Get Update(framesPassed)
        m_Frames = m_Frames - framesPassed
        Update = m_Frames
    End Property

    Public Sub NextFrame()
        m_currentIdx = m_currentIdx + 1
    End Sub

    Public Sub ResetInterval()

        m_Frames = m_updateInterval
        Exit Sub

        If Not IsNull(m_sequence) And UBound(m_sequence) > 1 Then

        'For i = 0 To totalSteps - 1
        '    currentStep = i
        '    duration = 20 ' Base duration of 20ms
            'Debug.print("TotalSteps: " & UBound(m_sequence)-1)
            Dim easeAmount : easeAmount = Round(m_currentIdx / UBound(m_sequence), 2) ' Normalize current step
            if easeAmount < 0 then
                easeAmount = 0
            elseif easeAmount > 1 then
                easeAmount = 1
            end if
            'Debug.print("Step: " & m_currentIdx)
            'Debug.print("Ease Amount: "& easeAmount)
            Dim newDuration : newDuration = 100 - Lerp(20, 80, EaseIn(easeAmount) )' Apply EaseInOut to duration
            Debug.print("Duration: "& Round(newDuration))
            'Dim newDuration : newDuration = 100- Lerp(20, 80, Spike(easeAmount) )' Apply EaseInOut to duration
            
            m_frames = newDuration
        Else
            m_Frames = m_updateInterval
        End If
    End Sub

End Class

Class LCSeqRunner
	
	Private m_name, m_items,m_currentItemIdx

    Public Property Get Name()
        Name=m_name
    End Property
    
	Public Property Let Name(input)
		m_name = input
	End Property

    Public Property Get CurrentItem()
        Dim items: items = m_items.Items()
        If m_currentItemIdx > UBound(items) Then
            m_currentItemIdx = 0
        End If
        If UBound(items) = -1 Then       
            CurrentItem  = Null
        Else
            Set CurrentItem = items(m_currentItemIdx)                
        End If
    End Property

    Private Sub Class_Initialize()    
        Set m_items = CreateObject("Scripting.Dictionary")
        m_currentItemIdx = 0
    End Sub

    Public Sub AddItem(item)
        If Not IsNull(item) Then
            If Not m_items.Exists(item.Name) Then
                    m_items.Add item.Name, item
            End If
        End If
    End Sub

    Public Sub RemoveAll()
        m_items.RemoveAll
    End Sub

    Public Sub RemoveItem(item)
        If Not IsNull(item) Then
            If m_items.Exists(item.Name) Then
                    m_items.Remove item.Name
            End If
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        If items(m_currentItemIdx).Repeat = False Then
            RemoveItem(items(m_currentItemIdx))
        End If
        m_currentItemIdx = m_currentItemIdx + 1
        If m_currentItemIdx > UBound(items) Then   
            m_currentItemIdx = 0    
        End If
    End Sub

    Public Function HasSeq(name)
        If m_items.Exists(name) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class


Function Lerp(startValue, endValue, amount)
    Lerp = startValue + (endValue - startValue) * amount
End Function

Function Flip(x)
    Flip = 1 - x
End Function

Function EaseIn(amount)
    EaseIn = amount * amount
End Function

Function EaseOut(amount)
    EaseOut = Flip(Sqr(Flip(amount)))
End Function

Function EaseInOut(amount)
    EaseInOut = Lerp(EaseIn(amount), EaseOut(amount), amount)
End Function

Function Spike(t)
    If t <= 0.5 Then
        Spike = EaseIn(t / 0.5)
    Else
        Spike = EaseIn(Flip(t)/0.5)
    End If
End Function