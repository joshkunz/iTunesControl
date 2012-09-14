COM_Init()
COM_Error(0) ;Un-Comment to see COM Errors

iTunes := COM_CreateObject("iTunes.Application")

/*
This is a little script to control iTunes Globaly

Info:
This Script requires the COM interface for AHK. It's included in all new versions of AHK.
If you are still holding on to an old version the COM package can be downloaded from
here: http://www.autohotkey.com/forum/topic22923.html

This will not work with the very old versions of iTunes
This Has been tested and working on iTunes 10

Instructions:

Quick Note On the Following
	-{Enter} is pressing the enter Key
	-You dont have to hold down the Ctrl+* key combo while entering data in brackets

Mouse Commands:
Left Click: Pause/Pause Music
Left Click+Drag-Right: Next Track
Left Click+Drag-Left: Previous Track

Keyboard Commands:
Ctrl+p: Play/Pause Music
Ctrl+RightButton: Next Track
Ctrl+LeftButton: PreviousTrack
Ctrl+UpButton: Volume Up
Ctrl+DownButton: Volume Down
Ctrl+s: Toggle Shuffle on/off
Ctrl+r+[1]: Set Repeat One
Ctrl+r+[a]: Set Repeat All
Ctrl+r+[n]: Set Repeat None
Ctrl+l+[PlaylistName]+{Enter}: Plays Playlist [PlaylistName] if it exsits *Playlist Name is Case-Sensitive
*/

;below are the Mouse Commands

^LButton::
MouseGetPos mx1, my1
sleep 100
MouseGetPos mx2, my2
changeX := mx2-mx1
changeY := my2-my1
if (changeX = 0 and changeY = 0) {
	COM_Invoke(iTunes, "PlayPause")
	} Else { 
		if (Abs(changeX) > Abs(changeY)) { ; Check to see if the mouse moved more horizontaly or verticaly
			if (changeX > 0) {
				;This Command is run when the mouse is dragged to the right
				COM_Invoke(iTunes, "NextTrack")
			} Else { 
			if (changeX < 0) {
				;This Command is run when the mouse is dragged to the left
				COM_Invoke(iTunes, "PreviousTrack")
				}
			}
	} Else { 
	if (Abs(changeY) > Abs(changeX)) {
		if (changeY > 0) {
				;This Command is run when the mouse is dragged up
			} Else { 
			if (changeY < 0) {
				;This Command is run when the mouse is dragged down
				}
		}
	}
}}
Return

;Below are the Keyboard Shortcuts

^p::COM_invoke(iTunes, "PlayPause") ;Toggle Play/Pause
^Right::COM_Invoke(iTunes, "NextTrack") 
^Left::COM_Invoke(iTunes, "PreviousTrack") 

^Delete::
currentSong := COM_Invoke(iTunes, "CurrentTrack")
COM_Invoke(currentSong, "Delete")
Return

^Up:: ;Turns Up the Volume
currentVolume := COM_Invoke(iTunes, "SoundVolume")
newVolume := currentVolume+5
COM_Invoke(iTunes, "SoundVolume", newVolume)
return

^Down:: ;Turns Down The volume
currentVolume := COM_Invoke(iTunes, "SoundVolume")
newVolume := currentVolume-5
COM_Invoke(iTunes, "SoundVolume", newVolume)
return

^s:: ;Toggles Shuffle
curPlaylist := COM_Invoke(iTunes, "CurrentPlaylist")
if (COM_Invoke(curPlaylist, "Shuffle") = -1) {
	COM_Invoke(curPlaylist, "Shuffle", 0)
} Else {
	COM_Invoke(curPlaylist, "Shuffle", -1)
	}
Return

^r:: ;Changes Repeat State 
curPlaylist := COM_Invoke(iTunes, "CurrentPlaylist")
Input, repeatWhat, L1
if repeatWhat = 1
	COM_Invoke(curPlaylist, "SongRepeat", 1)
else if repeatWhat = a
	COM_Invoke(curPlaylist, "SongRepeat", 2)
else if repeatWhat = n
	COM_Invoke(curPlaylist, "SongRepeat", 0)
	
^l:: ; Plays Playlist playlistName
Input, playlistName, C I, {Enter}
libs := COM_Invoke(iTunes, "LibrarySource")
playLists := COM_Invoke(libs, "Playlists")
playlist := COM_Invoke(playLists, "ItemByName", playlistName)
COM_Invoke(playlist, "PlayFirstTrack")
Return

^q::ExitApp 0 
