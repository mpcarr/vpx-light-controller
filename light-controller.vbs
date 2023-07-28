'***********************************************************************************************************************
' Lights State Controller - 8.0.0
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************

Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqOverride, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, useVpxLights

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
        useVpxLights = False
    End Sub

    Private Sub AssignStateForFrame(key, state)
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

    Public Sub RegisterLights(mode)

        Dim idx,tmp,vpxLight,lcItem
        If mode = "Lampz" Then
            
            For idx = 0 to UBound(Lampz.obj)
                If Lampz.IsLight(idx) Then
                    Set lcItem = new LCItem
                    If IsArray(Lampz.obj(idx)) Then
                        tmp = Lampz.obj(idx)
                        Set vpxLight = tmp(0)
                    Else
                        Set vpxLight = Lampz.obj(idx)
                        
                    End If
                    Lampz.Modulate(idx) = 1/100
                    Lampz.FadeSpeedUp(idx) = 100/30 : Lampz.FadeSpeedDown(idx) = 100/120
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next        
        ElseIf mode = "VPX" Then
            useVpxLights = True
            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next  
        End If
    End Sub

	Public Sub AddLight(light, idx)
        If m_lights.Exists(light.name) Then
            Exit Sub
        End If
        Dim lcItem : Set lcItem = new LCItem
        lcItem.Init idx, light.BlinkInterval, Array(light.color, light.colorFull), light.name, light.x, light.y
        m_lights.Add light.Name, lcItem
        m_seqRunners.Add "lSeqRunner" & CStr(light.name), new LCSeqRunner
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
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1)
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
            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
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

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
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

    Public Sub UpdateBlinkInterval(light, interval)
        If m_lights.Exists(light.name) Then
            light.BlinkInterval = interval
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs.Item(light.name & "Blink").UpdateInterval = interval
            End If
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
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount)
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
            m_pulse.Add name, (new PulseState)(m_lights(name), profile, 0, m_pulseInterval, repeatCount)
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
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
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
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverride.AddItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(lcSeq)
        m_seqOverride.RemoveItem lcSeq
        If IsNull(m_seqOverride.CurrentItem) Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        m_seqOverride.RemoveAll
        Dim light
		For Each light in m_lights.Keys()
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
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
        If useVpxLights = True Then
            Exit Sub
        End If

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

    Private Function GetTmpLight(idx)
        If useVpxLights = True Then
          If IsArray(Lights(idx) ) Then	'if array
                Set GetTmpLight = Lights(idx)(0)
            Else
                Set GetTmpLight = Lights(idx)
            End If
        Else
            If IsArray(Lampz.obj(idx) ) Then	'if array
                Set GetTmpLight = Lampz.obj(idx)(0)
            Else
                Set GetTmpLight = Lampz.obj(idx)
            End If
        End If
        
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
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
                    AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).light.Color, m_pulse(lightKey).light.Idx)
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey).interval - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey).idx
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If
                    
                    Dim pulses : pulses = m_pulse(lightKey).pulses
					Dim pulseCount : pulseCount = m_pulse(lightKey).Cnt
                    If pulseIdx > UBound(m_pulse(lightKey).pulses) Then
						m_pulse.Remove lightKey    
						If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                        End If
                    Else
						m_pulse.Remove lightKey
                        m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount)
                    End If
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(0, Null, lcItem.Idx)
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
                If Not IsNull(m_vpxLightSyncCollection) Then
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

                            'TODO - Fix VPX Fade
                            If Not IsNull(m_tableSeqFadeUp) Then
                                Lampz.FadeSpeedUp(syncLight.Idx) = m_tableSeqFadeUp
                            End If
                            If Not IsNull(m_tableSeqFadeDown) Then
                                Lampz.FadeSpeedDown(syncLight.Idx) = m_tableSeqFadeDown
                            End If
                    
                            AssignStateForFrame syncLight.name, (new FrameState)(lx.GetInPlayState*100,color, syncLight.Idx)                     
                        End If
                    Next
		        End If
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
                            AssignStateForFrame syncClearLight.name, (new FrameState)(0, Null, syncClearLight.idx) 
                            'TODO - Only do fade speed for lampz
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
                Dim idx : idx = m_currentFrameState(frameStateKey).idx
                'Debug.print("Changing light idx: " & CStr(m_currentFrameState(frameStateKey).idx) & " -> " & CStr(m_currentFrameState(frameStateKey).level) & ". FrameTime: " & m_frametime)
                
                Dim newColor : newColor = m_currentFrameState(frameStateKey).colors
                Dim bUpdate

                If Not IsNull(newColor) Then
                    'Debug.Print("Updating color")
                    'Check current color is the new color coming in, if not, set the new color.
                    
                    Set tmpLight = GetTmpLight(idx)

					Dim c, cf
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
            	End If

                If useVpxLights = False Then
                    If bUpdate Then
                        'Update lamp color
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
                        If Lampz.UseCallBack(idx) then Proc Lampz.name & idx,Lampz.Lvl(idx)*Lampz.Modulate(idx)	'Force Callbacks Proc
                    End If
                    Lampz.state(idx) = CInt(m_currentFrameState(frameStateKey).level) 'Lampz will handle redundant updates
                Else
                    If IsArray(Lights(idx)) Then
                        For Each x in Lights(idx)
                            If bUpdate Then 
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                            End If
                            x.State = m_currentFrameState(frameStateKey).level/100
                        Next
                    Else
                        If bUpdate Then 
                            If Not IsNull(c) Then
                                Lights(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lights(idx).colorFull = cf
                            End If
                        End If
                        Lights(idx).State = m_currentFrameState(frameStateKey).level/100
                    End If
                End If


           
                
				 
            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HexToInt(hex)
        HexToInt = CInt("&H" & hex)
    End Function

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
        dim lsName, isSeqEnd
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            isSeqEnd = True
        Else
            isSeqEnd = False
        End If

        dim lightInSeq
        For each lightInSeq in lcSeq.LightsInSeq
            If isSeqEnd Then
                
            'Needs a guard here for something, but i've forgotten. 
            'I remember: Only reset the light if there isn't frame data for the light. 
            'e.g. a previous seq has affected the light, we don't want to clear that here on this frame
                If m_lights.Exists(lightInSeq) = True AND NOT m_currentFrameState.Exists(lightInSeq) Then
                   AssignStateForFrame lightInSeq, (new FrameState)(0, Null, m_lights(lightInSeq).Idx)
                End If
            Else
                '
                If m_currentFrameState.Exists(lightInSeq) Then
                    'already frame data for this light.
                    'replace with the last known state from this seq
                    If Not IsNull(lcSeq.LastLightState(lightInSeq)) Then
						AssignStateForFrame lightInSeq, lcSeq.LastLightState(lightInSeq)
                    End If
                End If

            End If
        Next

        If isSeqEnd Then
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        End If

        If Not IsNull(seqRunner.CurrentItem) Then
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
							If lsName(2) = "FFFFFF" Then
                                AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                            Else
                                AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                            End If
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        End If
						lcSeq.LastLightState(name) = m_currentFrameState(name)
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
                        If lsName(2) = "FFFFFF" Then
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        End If
                    Else
                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    End If
                    lcSeq.LastLightState(name) = m_currentFrameState(name)
                End If
            End If

            '

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

Class FrameState
    Private m_level, m_colors, m_idx

    Public Property Get Level(): Level = m_level: End Property
    Public Property Let Level(input): m_level = input: End Property

    Public Property Get Colors(): Colors = m_colors: End Property
    Public Property Let Colors(input): m_colors = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public default function init(level, colors, idx)
		m_level = level
		m_colors = colors
		m_idx = idx 

		Set Init = Me
    End Function

    Public Function ColorAt(idx)
        ColorAt = m_colors(idx) 
    End Function
End Class
 
Class PulseState
    Private m_light, m_pulses, m_idx, m_interval, m_cnt

    Public Property Get Light(): Set Light = m_light: End Property
    Public Property Let Light(input): Set m_light = input: End Property

    Public Property Get Pulses(): Pulses = m_pulses: End Property
    Public Property Let Pulses(input): m_pulses = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public Property Get Interval(): Interval = m_interval: End Property
    Public Property Let Interval(input): m_interval = input: End Property

    Public Property Get Cnt(): Cnt = m_cnt: End Property
    Public Property Let Cnt(input): m_cnt = input: End Property

    Public default function init(light, pulses, idx, interval, cnt)
		Set m_light = light
		m_pulses = pulses
		'debug.Print(Join(Pulses))
		m_idx = idx 
		m_interval = interval
		m_cnt = cnt

		Set Init = Me
    End Function

    Public Function PulseAt(idx)
        PulseAt = m_pulses(idx) 
    End Function
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
	
	Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
		m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
	Public Property Let Sequence(input)
		m_sequence = input
        dim item, light, lightItem
        for each item in input
            If IsArray(item) Then
                for each light in item
                    lightItem = Split(light,"|")
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If    
                next
            Else
                lightItem = Split(item,"|")
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
            End If
        next
	End Property

    Public Property Get LastLightState(light)
		If m_lastLightStates.Exists(light) Then
			dim c : Set c = m_lastLightStates(light)
			Set LastLightState = c
		Else
			LastLightState = Null
		End If
    End Property

    Public Property Let LastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
		If input.level > 0 Then
			m_lastLightStates.Add light, input
		End If
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
        m_Frames = input
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
        Set m_lightsInSeq = CreateObject("Scripting.Dictionary")
        Set m_lastLightStates = CreateObject("Scripting.Dictionary")
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
            'Debug.print("Duration: "& Round(newDuration))
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
        Dim item
        For Each item in m_items.Keys()
            m_items(item).ResetInterval
            m_items(item).CurrentIdx = 0
            m_items.Remove item
        Next
    End Sub

    Public Sub RemoveItem(item)
        If Not IsNull(item) Then
            If m_items.Exists(item.Name) Then
                    item.ResetInterval
                    item.CurrentIdx = 0
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