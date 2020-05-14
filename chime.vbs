Option Explicit

' Change the path and filename below
Dim sound 

sound = "./media/" & hour(now()) & ".mp3"

' WScript.Echo sound


Dim o : Set o = CreateObject("wmplayer.ocx")
With o
    .url = sound
    .controls.play
    While .playstate <> 1
        wscript.sleep 100
    Wend
    .close
End With

Set o = Nothing
