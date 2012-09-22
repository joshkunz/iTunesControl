iTunesControl
=====

Global hotkeys for iTunes on windows.

## What's in here?

* __iTunesControl.py__ - The original python script, requires pyHook and the python
win32 hooks installed. It's probably working.
* __iControl.ahk__ - Autohotkey re-write of the original script. Mostly just for noodling
around with autohotkey.
* __iTunesCOM__ - The ridiculously hard to find iTunes COM interface documentation. I've
put it up on github pages [here](joshkunz.github.com/iTunesControl).

## Hotkeys

### Basic Keyboard Shortcuts

<table>
    <tr>
        <td>Alt + Q</td>
        <td>Qut the application</td>
    <tr>
        <td>Alt + P</td>
        <td>Toggle Play/Pause</td>
    </tr>
    <tr>
        <td>Alt + Right</td>
        <td>Next Track</td>
    </tr>
    <tr>
        <td>Alt + Left</td>
        <td>Previous Track</td>
    </tr>
    <tr>
        <td>Alt + Up</td>
        <td>Volume Up</td>
    </tr>
    <tr>
        <td>Alt + Down</td>
        <td>Volume Down</td>
    </tr>
    <tr>
        <td>Alt + (-)</td>
        <td>Mute (Set volume to 0)</td>
    </tr>
    <tr>
        <td>Alt + (+)</td>
        <td>Max Volume (Set volume to 100)</td>
    </tr>
    <tr>
        <td>Alt + S</td>
        <td>Toggle Shuffle</td>
    </tr>
</table>

### Advanced Keyboard Shortcuts

Note: the `Alt` key must be held down durning the 
entire command.

<table>
    <tr>
        <td>Alt + R + {1, A, N}</td>
        <td>Toggle repeat. R+1 is repeat one. R+A is repeat
        all. R+N is repeat none.</td>
    </tr>
    <tr>
        <td>Alt + H + [Playlist Name] + Enter</td>
        <td>Play playlist [Playlist Name] (Alt key must be held
        down while entering playlist name</td>
    </tr>
    <tr>
        <td>Alt + O + {1, 2, 3} + [Search Term] + Enter</td>
        <td>Search for term [Search Term] based on {1, 2, 3},
        put those songs into a new playlist, and start playing
        the playlist. 1: Search by artist name. 2: Search by
        album name. 3: Search by song name. (Again, the Alt key
        needs to be held down during the entire command)</td>
    </tr>
</table>

### Mouse Commands

<table>
    <tr>
        <td>Alt + Click</td>
        <td>Toggle Play/Pause</td>
    </tr>
    <tr>
        <td>Alt + Click and Drag to the Right</td>
        <td>Next track</td>
    </tr>
    <tr>
        <td>Alt + Click and Drag to the Left</td>
        <td>Previous track</td>
    </tr>
    <tr>
        <td>Alt + Scroll Up</td>
        <td>Volume Up</td>
    </tr>
    <tr>
        <td>Alt + Scroll Down</td>
        <td>Volume Down</td>
    </tr>
</table>

### Some Tips and Cautions

* Alt + H + "Music" - Will play the playlist "music" which is the
entire library. It's good for switching back once you get bored of
a playlist.
* If for some reason you want to used in a program that has
been "run as administrator" this program also needs to be run as
administrator.

## Design and Hacking (really it's an apology)

I wrote this program back when I was about 15, and really new to python.
I think this is the first program I wrote that used threads, and I don't
think I quite knew what that meant besides "you can run more than one
thing at once" and "you have to pass messages with a queue or it will
crash". I've tried to hold off on releasing this because it's in such 
bad shape, but I haven't had the time or the patience to clean it
up myself. Good luck if you try, but be sure to open a pull request
with improvements. Since the thing has about zero comments explaining
what it does, I'll add a brief description of the design below.

Basically it works like this: pyHook, basically keylogs every keystroke,
and passes it back to the operating system. If a keystroke is entered
and the Alt key is depressed at the same time, the program will
block the keystroke form reaching the other applications, and put it
into a queue. (There a two seperate queues, one for the mouse and another
for the keyboard). There are two threads that are basically loops waiting
for an item to come through the queue, both of the threads maintain a 
connection to iTunes using it's COM interface (a little like dbus or
applescript for the uninitiated). Once an item is received, they have
a huge conditional sections that basically executes COM statements
based on the key pressed and that's it. Things get a little more complicated
with the text-entry stuff, basically it just bypasses the other conditionals.

Have fun, and good luck using it. (and please believe me when I say I write
much better code these days.)
