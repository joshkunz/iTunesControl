import win32com.client
if win32com.client.gencache.is_readonly == True:

    win32com.client.gencache.is_readonly = False

    win32com.client.gencache.Rebuild()
    
from win32com.client.gencache import EnsureDispatch
import pythoncom
import pyHook
import sys
from time import sleep
import threading
import Queue
import multiprocessing


global Kque
global Mque
global sendInfo
global LastAlt
LastAlt = 0
sendInfo = True
Kque = Queue.Queue()
Mque = Queue.Queue()	
	

class Actor(threading.Thread):
	def run(self):
		global Kque
		global LastAlt
		global sendInfo
		pythoncom.CoInitialize()
		iTunes = EnsureDispatch("iTunes.Application")
		self.To_Delete = []
		print "Ready"
		while 1:
			command = Kque.get()
			
			if command[2] > 0:
				LastAlt = command[0]
				sendInfo = False
					
			if command[0]-LastAlt > 200:
				sendInfo = True
				
			try:
				if command[1] == "P" and command[2] > 0:
					iTunes.PlayPause()
				elif command[1] == "Right" and command[2] > 0:
					iTunes.NextTrack()
				elif command[1] == "Left" and command[2] > 0:
					iTunes.BackTrack()
				elif command[1] == "Up" and command[2] > 0:
					iTunes.SoundVolume += 5
				elif command[1] == "Down" and command[2] > 0:
					iTunes.SoundVolume -= 5
				elif command[1] == "Oem_Minus" and command[2] > 0:
					iTunes.SoundVolume = 0
				elif command[1] == "Oem_Plus" and command[2] > 0:
					iTunes.SoundVolume = 100
				elif command[1] == "S" and command[2] > 0:
					MainPlaylist = iTunes.CurrentPlaylist
					if MainPlaylist.Shuffle == 1:
						MainPlaylist.Shuffle = 0
					elif MainPlaylist.Shuffle == 0:
						MainPlaylist.Shuffle = 1
					else:
						pass
					
				elif command[1] == "Finish" and command[2] > 0:
					while len(self.To_Delete) > 0:
						temp_l = iTunes.LibrarySource.Playlists.ItemByName(self.To_Delete.pop())
						temp_l.Delete()
				elif command[1] == "R" and command[2] > 0:
					MainPlaylist = iTunes.CurrentPlaylist
					Kque.task_done()
					repeat = Kque.get()
					if repeat[1] == "1" and repeat[2] > 0:
						MainPlaylist.SongRepeat = 1
					elif repeat[1] == "A" and repeat[2] > 0:
						MainPlaylist.SongRepeat = 2
					elif repeat[1] == "N" and repeat[2] > 0:
						MainPlaylist.SongRepeat = 0
					else:
						pass
				elif command[1] == "H" and command[2] > 0:
					print "Enter Playlist Name:"
					char_list = []
					Kque.task_done()
					pressed_key = Kque.get()
					
					while pressed_key[2] > 0:
						char_list.append(pressed_key[1])
						Kque.task_done()
						pressed_key = Kque.get()
						
					ret_string = ""
					Caps = False
					Shift = False
					for x in char_list:
						val = x.lower()
						if val not in ["space", "lshift", "rshift", "capital"]:
							if Shift == True:
								val =  val.upper()
								Shift = False
							elif Caps == True:
								val = val.upper()
							else:
								pass
							ret_string += val
							
						elif val == "space":
							ret_string += " "
							
						elif val in ["lshift", "rshift"]:
							Shift = True
							
						elif val == "capital":
							if Caps == True:
								Caps = False
							elif Caps == False:
								Caps = True
							else:
								pass
					try:
						gotoPlaylist = iTunes.LibrarySource.Playlists.ItemByName(ret_string)
						gotoPlaylist.PlayFirstTrack()
						print "Playing Playlist: %s"% ret_string
					except:
						print "Playlist %s Not Found"% ret_string
						
				elif command[1] == "O" and command[2] > 0:
				
					Kque.task_done()
					repeat = Kque.get()
					Op = None
					if repeat[1] == "1" and repeat[2] > 0:
						Op = "1"
					elif repeat[1] == "2" and repeat[2] > 0:
						Op = "2"
					elif repeat[1] == "3" and repeat[2] > 0:
						Op = "3"
					else:
						pass
					
					print "Enter Char String"
					char_list = []
					Kque.task_done()
					pressed_key = Kque.get()
					
					while pressed_key[2] > 0:
						char_list.append(pressed_key[1])
						Kque.task_done()
						pressed_key = Kque.get()
						
					ret_string = ""
					Caps = False
					Shift = False
					for x in char_list:
						val = x.lower()
						if val not in ["space", "lshift", "rshift", "capital"]:
							if Shift == True:
								val =  val.upper()
								Shift = False
							elif Caps == True:
								val = val.upper()
							else:
								pass
							ret_string += val
							
						elif val == "space":
							ret_string += " "
							
						elif val in ["lshift", "rshift"]:
							Shift = True
							
						elif val == "capital":
							if Caps == True:
								Caps = False
							elif Caps == False:
								Caps = True
							else:
								pass
							
					Liby = iTunes.LibraryPlaylist
					Tracks = Liby.Tracks
					if Op == "1":
						print "Scaning for artist: %s"% ret_string
						track_list = []
						for track in Tracks:
							if track.Artist.lower() == ret_string.lower():
								track_list.append(track)
					
					elif Op == "2":
						print "Scaning for album: %s"% ret_string
						track_list = []
						for track in Tracks:
							if track.Album.lower() == ret_string.lower():
								track_list.append(track)
								
					elif Op == "3":
						print "Scaning for Song Name: %s"% ret_string
						track_list = []
						for track in Tracks:
							if track.Name.lower() == ret_string.lower():
								track_list.append(track)
								
					else:
						pass
							
					if len(track_list) > 0: 
						temp_list = iTunes.CreatePlaylist(ret_string)
						self.To_Delete.append(ret_string)
						temp_list = win32com.client.CastTo(temp_list, 'IITUserPlaylist')
						for track in track_list:
							temp_list.AddTrack(track)
							
						temp_list.PlayFirstTrack()
						print "Done"
					else:
						print "No Tracks Found"
								
						
					
				else:
					pass
			except pythoncom.com_error, e:
				print e
			Kque.task_done()
			
class Actor2(threading.Thread):
	def run(self):
		global LastAlt
		global Mque
		global sendInfo
		pythoncom.CoInitialize()
		iTunes = EnsureDispatch("iTunes.Application")
		self.RMouseDown = False
		self.PosStart = None
		self.PosEnd = None
		print "Ready"
		while 1:
			command = Mque.get()
			
			if sendInfo == False and command[1] == 513:
				self.RMouseDown = True
				
			if sendInfo == False and self.RMouseDown == True and self.PosStart == None and command[1] == 512:
				self.PosStart = command[2]
				
			if sendInfo == False and self.RMouseDown == True and command[1] == 512:
				self.PosEnd = command[2]
				
			try:
				if sendInfo == False and self.RMouseDown == True and command[1] == 514:
					self.RMouseDown = False
					if self.PosStart != None and self.PosEnd != None:
						if self.PosStart[0] < self.PosEnd[0]:
							iTunes.NextTrack()
						elif self.PosStart[0] > self.PosEnd[0]:
							iTunes.BackTrack()
						else:
							pass
					else:
						iTunes.PlayPause()
					self.PosStart = None
					self.PosEnd = None
					
				if sendInfo == False and command[3] != 0:
					if command[3] > 0:
						iTunes.SoundVolume += 2
					elif command[3] < 0:
						iTunes.SoundVolume -= 2
					else:
						pass
			except pythoncom.com_error, e:
				print e
							
			Mque.task_done()
			
			
	

thread = Actor2()
thread.setDaemon(True)
thread.start()

thread = Actor()
thread.setDaemon(True)
thread.start()

		
def OnKeyboardEvent(event):
	global Kque
	global sendInfo
	
	if event.Key == "Q" and event.Alt > 0:
		Kque.put((0, "Finish", 32))
		while len(thread.To_Delete) > 0:
			sleep(0.2)
		print "Thanks!"
		sys.exit(0)

	Kque.put([event.Time, event.Key, event.Alt])

	if sendInfo == True:
		return True
	else:
		return False

def OnMouseEvent(event):
	global Mque
	global LastAlt
	global sendInfo
	# called when mouse events are received
	
	Mque.put([event.Time, event.Message, event.Position, event.Wheel])
	

	if sendInfo != True:
		if event.Message == 513 or event.Message == 514:
			if event.Time-LastAlt > 150:
				sendInfo = True
				return True
			else:
				return False
		elif event.Message == 522:
			if event.Time-LastAlt > 150:
				sendInfo = True
				return True
			else:
				return False
		else:
			return True
	else:
		return True
	
# create a hook manager
hm = pyHook.HookManager()
# watch for all mouse events
hm.KeyDown = OnKeyboardEvent
		# set the hook
hm.HookKeyboard()

hm.MouseAll = OnMouseEvent
# set the hook
hm.HookMouse()
# wait forever
pythoncom.PumpMessages()

if __name__ == '__main__':
	print "paapy nama"
