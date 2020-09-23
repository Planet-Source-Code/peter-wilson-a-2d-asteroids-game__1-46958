Asteroids Game
==============
"Explosive rapid-fire space action!

You have been sent to outer space in a new space ship armed with a powerful bullet space weapon to shoot down the asteroids that are an ever growing threat. Your objective is to destroy all the asteroids, and to not get destroyed. Be careful, and good luck."

This game was written using Visual Basic. The source code for this game can be downloaded from http://dev.midar.com/
Copyright © 2003 Peter Wilson - All Rights Reserved.


Peter's Tip
===========
Turn off all the lights and play in the dark. Adjust your computer monitors contrast control to maximum, and then slowly adjust the monitors brightness control until the backgound turns inky-black. I've also drawn some cross-hairs onto the screen which are very dim. They should be barely visible when playing under the correct conditions. Don't forget to change the Options under Init_Game.


Keyboard Controls
=================
ESC		:	Quit game. Please use this to avoid problems shutting down the MIDI driver.
Arrows		:	Rotate the Player space ship, Thrust forwards, and Dead-stop.
Left-Control	:	Fire Primary Weapon.
Tab		:	Overview Map (Toggle)
Shift		:	Apply a scale matrix to the geometry to the game (ie. Magnify objects temporarily)
Space Bar	:	Change Levels, increase number of Asteroids.
C Key		:	Do not clear the screen - cool effects.
P Key		:	Pause Button (Toggle). Note: The timer control keeps going... I will put this to good use later.
Mouse-Down	:	Temporarily freeze the game (similar to the Pause button)


Features to look out for...
===========================
* There are not too many comments at this stage, however when I finish the game, then I will fully document it.
  If you have any questions, just ask and I will explain.

* Lots of cool stuff, scattered throughout the code!
  I personally love the routine I wrote to create a random asteroid.... this is really what started the whole project.
  (It's really just a distorted circle!)

* Matrix Concatenation using Matrix Multiplication
  (Note: The order in which the Matrices are multiplied together.)

	matResult = MatrixIdentity
	matResult = MatrixMultiply(matResult, m_matScale)
	matResult = MatrixMultiply(matResult, matRotationAboutZ)
	matResult = MatrixMultiply(matResult, matTranslate)
	matResult = MatrixMultiply(matResult, m_matViewMapping)

* Matrix * Vector multiplication
  (This is the fun part, that changes our 3D vector into 2D screen space)

	For intJ = LBound(.Vertex) To UBound(.Vertex)
		.TVertex(intJ) = MatrixMultiplyVector(matResult, .Vertex(intJ))
	Next intJ


Version History
===============
July 2003	: * Initial released. Code not fully functional. Can't shoot Asteroids.
		  * Make sure you quit the game using the ESC key, otherwise the MIDI driver will not shut
		    down correctly, and you'll have to close down VB to reset it.
		  * There is a small bug when shooting the player ammo. Extra explosions appear in the centre of the window.
		  * To make the player move faster or slower, change the code in the Keyboard module.
		  * Disable MIDI to speed up game (a little bit)
		  * Take the MIDI code out completely (ie. delete it, or conditionally compile it) to really speed up this game!

July 2003 (b)	: * Added a crude collision detection system.
		  * Player's Ship looks really cool when it blows up.
		  * Fixed Zooming to make it smoother.

Feedback
========
If you've got any questions, praise or comments then send me an e-mail.
If you feel so inclined to vote for this code on Planet-Source-Code, then that would be good too although not necessary.


Peter Wilson
peter@midar.com
http://dev.midar.com/Tutorials/Computers/Programming/VB/VBLessons.asp
--
end of line
--
