' =============================================================================================================================
' fadeVolume plugin.  Gradually changes volume of an audio output connector over time in response to UDP message.
' v1.2 - non-blocking fade (roTimer, re-armed each tick), bug fixes, renamed from fade_volume. Based on v1.1 (18 July 2025).
'
' HOW TO USE:	Send UDP message on port 6000 "fadeVolume!<targetVolume>!<durationMs>" (e.g. "fadeVolume!50!5000")
' REQUIREMENTS:	Presentation must have two user variables:
' 1: "storedVolume", default value of 100.
' 2: "audioDevice", with default of one of: "hdmi", "usb", "spdif", "analog"
'
' KNOWN ISSUE: plugin will also respond on UDP Receiver port set in Presentation Settings / Interactive / Networking
' =============================================================================================================================


Function fadeVolume_Initialize(msgPort As Object, userVariables As Object, bsp As Object) As Object
	print "+++ fadeVolume_Initialize - plugin entry"
	return newfadeVolume(msgPort, userVariables, bsp)
End Function


Function newfadeVolume(msgPort As Object, userVariables As Object, bsp As Object) As Object
	print "+++ newfadeVolume - Initializing plugin object"

	s = {}
	s.msgPort = msgPort
	s.userVariables = userVariables
	s.bsp = bsp
	s.ProcessEvent = fadeVolume_ProcessEvent
	s.objectName = "fadeVolume_object"
	s.debug = false				' set true for verbose variable logging

	s.audioOutput = invalid		' created lazily on first fade
	s.fade = invalid			' state of the fade in progress (invalid when idle)
	s.stepIntervalMs = 50		' how often the volume is updated during a fade

	' Timer that drives the fade without blocking the event loop
	s.fadeTimer = CreateObject("roTimer")
	s.fadeTimer.SetPort(msgPort)
	s.fadeTimer.SetElapsed(0, s.stepIntervalMs)

	' UDP setup
	s.udpReceiverPort = 6000
	s.udpReceiver = CreateObject("roDatagramReceiver", s.udpReceiverPort)
	if s.udpReceiver = invalid then
		print "+++ Error: could not open UDP port"; s.udpReceiverPort
	else
		s.udpReceiver.SetPort(msgPort)
	end if

	' Copy variable defaults to current values
	InitUserVar(s, "audioDevice")
	InitUserVar(s, "storedVolume")
	s.currentVolume = ReadStoredVolume(s)

	return s
End Function


Function fadeVolume_ProcessEvent(event As Object) As Boolean
	' NB: called as a method, so "m" is the plugin object here
	s = m

	if type(event) = "roDatagramEvent" then
		msg = lcase(event)
		if left(msg, 10) = "fadevolume" then
			print "+++ Received UDP Message: "; msg
			ParseFadeMsg(s, msg)
			return true
		end if
	else if type(event) = "roTimerEvent" then
		if event.GetSourceIdentity() = s.fadeTimer.GetIdentity() then
			FadeTick(s)
			return true
		end if
	end if

	return false
End Function


' Validate "fadeVolume!<volume>!<durationMs>" and start the fade
Sub ParseFadeMsg(s As Object, msg As String)
	fields = CreateObject("roRegex", "!", "").Split(msg)
	if fields.Count() <> 3 then
		print "+++ Incorrect number of fields for fadeVolume command: "; msg
		return
	end if

	isNumber = CreateObject("roRegex", "^[0-9]+$", "")
	volStr = fields[1].Trim()
	durStr = fields[2].Trim()
	if not isNumber.IsMatch(volStr) or not isNumber.IsMatch(durStr) then
		print "+++ Volume and duration must be whole numbers: "; msg
		return
	end if

	StartFade(s, val(volStr), val(durStr))
End Sub


Sub StartFade(s As Object, targetVolume As Integer, durationMs As Integer)
	targetVolume = ClampVolume(targetVolume)

	' A new command supersedes any fade in progress; carry on from the actual current volume
	CancelFade(s)

	if targetVolume = s.currentVolume then
		print "+++ Already at volume"; targetVolume; ". Nothing to do!"
		return
	end if

	if not InitializeAudioOutput(s) then return

	if durationMs <= 0 then
		SetOutputVolume(s, targetVolume)
		SetUserVar(s, "storedVolume", targetVolume)
		return
	end if

	s.fade = {
		startVolume: s.currentVolume,
		targetVolume: targetVolume,
		durationMs: durationMs,
		clock: CreateObject("roTimespan")
	}
	s.fade.clock.Mark()
	s.fadeTimer.Start()
	print "+++ Fading from"; s.currentVolume; " to"; targetVolume; " over"; durationMs; "ms"
End Sub


' Called on every timer tick while a fade is running.
' Volume is derived from elapsed time, so timer jitter can't make the fade run long or short.
Sub FadeTick(s As Object)
	f = s.fade
	if f = invalid then
		s.fadeTimer.Stop()
		return
	end if

	elapsed = f.clock.TotalMilliseconds()
	if elapsed >= f.durationMs then
		newVolume = f.targetVolume
	else
		newVolume = ClampVolume(int(f.startVolume + (f.targetVolume - f.startVolume) * elapsed / f.durationMs + 0.5))
	end if

	if newVolume <> s.currentVolume then SetOutputVolume(s, newVolume)

	if elapsed >= f.durationMs then
		s.fadeTimer.Stop()
		s.fade = invalid
		SetUserVar(s, "storedVolume", s.currentVolume)
	else
		' roTimer fires only once per Start() on this firmware, so re-arm for the next step
		s.fadeTimer.SetElapsed(0, s.stepIntervalMs)
		s.fadeTimer.Start()
	end if
End Sub


Sub CancelFade(s As Object)
	if s.fade <> invalid then
		s.fadeTimer.Stop()
		s.fade = invalid
		SetUserVar(s, "storedVolume", s.currentVolume)
	end if
End Sub


Function ClampVolume(v As Integer) As Integer
	if v > 100 then return 100
	if v < 0 then return 0
	return v
End Function


Sub SetOutputVolume(s As Object, volume As Integer)
	s.audioOutput.SetVolume(volume)
	s.currentVolume = volume
End Sub


Function InitializeAudioOutput(s As Object) As Boolean
	if s.audioOutput <> invalid then return true

	audDev = GetUserVar(s, "audioDevice")
	if audDev = invalid or audDev = "" then
		print "+++ Failed to retrieve audioDevice variable. Using 'analog' as fallback."
		audDev = "analog"
	end if

	s.audioOutput = CreateObject("roAudioOutput", audDev)
	if s.audioOutput = invalid then
		print "+++ Error: Failed to initialize audio output for: "; audDev
		return false
	end if

	print "+++ Audio output initialized: "; audDev
	return true
End Function


Function ReadStoredVolume(s As Object) As Integer
	v = GetUserVar(s, "storedVolume")
	if v = invalid then return 100
	return ClampVolume(int(val(v)))
End Function


Sub InitUserVar(s As Object, varName As String)
	if s.userVariables.DoesExist(varName) then
		varLookup = s.userVariables.Lookup(varName)
		defVal = varLookup.defaultvalue$
		if SetUserVar(s, varName, defVal) then
			print "+++ Variable '" + varName + "' initialized with default: "; defVal
		else
			print "+++ Variable '" + varName + "' initialize failed."
		end if
	else
		print "+++ Variable '" + varName + "' not found for initialize."
	end if
End Sub


' Returns the variable's current value as a string, or invalid on failure
Function GetUserVar(s As Object, varName As String) As Dynamic
	if s.userVariables.DoesExist(varName) then
		varLookup = s.userVariables.Lookup(varName)
		if varLookup <> invalid then
			currentValue = varLookup.GetCurrentValue()
			if s.debug then print "+++ Variable '" + varName + "' current value: "; currentValue
			return currentValue
		end if
	end if
	print "+++ Variable '" + varName + "' not found or invalid."
	return invalid
End Function


' Returns true on success. User variables hold strings, so numbers are converted.
Function SetUserVar(s As Object, varName As String, newValue As Dynamic) As Boolean
	if s.userVariables.DoesExist(varName) then
		varLookup = s.userVariables.Lookup(varName)
		if varLookup <> invalid then
			if type(newValue) <> "String" and type(newValue) <> "roString" then
				newValue = Str(newValue).Trim()
			end if
			varLookup.SetCurrentValue(newValue, true)
			if s.debug then print "+++ Variable '" + varName + "' set to: "; newValue
			return true
		end if
	end if
	print "+++ Variable '" + varName + "' not found or invalid."
	return false
End Function
