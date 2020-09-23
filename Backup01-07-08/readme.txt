This version adds an updated undo that has been expanded to cover any edit.
Also the user settings are saved to an initial file. The brush option now works.
And the code has been craped up a little more. Hope you enjoy it.


Updated Help and usage

When you run the program you will see two picture boxes and one frame. The picture box on the left will be referred to as the grid and the grid blocks will be referred to as pixels. The picture on the right will be the Source. The shape control on the source will be known as the "Selection Square". The frame can just be the Grid tools.

___________________________________________________________________________________________
The grid options are as follows:

 Clicking on the grid with the left mouse button will highlight that pixel if the "Enable Select Tool" is checked and color that pixel the currently selected color if it is not checked (Optionally you can select a drawing color by clicking on the pallet or by selecting the command "Color Chooser"). Dragging the mouse with "Enable select tool" unchecked will color all pixels that you move over.
 
 Clicking on the grid with the right mouse button will, if "Enable Select Tool" is unchecked, select that color under the mouse. If "Enable Select Tool" is checked the right mouse button will display a popup menu. Dragging the mouse with enable select tool checked will draw a focus rectangle and releasing the mouse will select all pixels within that area, and color that area otherwise

 The popup menu offers a few paste options. and a "Select Inside outline" which when selected will scan all the pixels in the grid in the top to bottom order if it finds a selected pixel it will continue selecting pixels in that line until it reaches the end or finds another selected pixel. Note this function needs some work; but, in essence if you select the top and bottom pixels of an area this function will select all the rest. The paste options are clear on what they do but a note on the usage in order. Once you have selected the pixels you want to copy. Right click on the source to move the selection square, Right click on the grid to display the menu and select the paste option you want. As long as you don’t left click in the source you can move and paste as many times as you want. NOTE: If you make a selection and move the selection square to a new position on the source you can't then make any other selections. There are no error handlers to verify this and it will not work. This is because when you start selecting an array is created with the bounds relative to the selection squares position on the source. Therefore when you move the selection square and start selecting pixels they fall outside those initial bounds.

 There is also another feature relative to the pixels being selected, and that is the ability to move pixels in the picture. When you select the "Enable Select Tool" a new frame will show four buttons to move with in addition you can use your arrow keys.

 And finally you can save these selected pixels to a file, reopen them and paste them any where in the picture the selection square is. Once opened into the app the pixel information is just loaded into the arrays but performing any action which would normally result in the pixels being written to the source like move or paste that is when they will be copied to the source image. All the same rules apply here as they do when selecting with the mouse.

 Just for fun I added a save grid as bitmap function I had a couple of pictures that looked good with a one pixel wide grid.

Update:
    There is a pattern brush feature that needs to be explained. To create a brush you need to select pixels along the x cord
The height you wish your brush to be. Just a single line will do as the brush will always be square, the width is determined
by the height. Those pixels that make up the square will be loaded into an array. Each time you click with the mouse or drag
The pattern you selected will be drawn to the screen using the pixel you select as the upper left. Dragging with the mouse produces
A tiled effect, top to bottom or side by side. You can save this brush pattern as a file and reload it any time.
_____________________________________________________________________________________________
The Source options

	Clicking the source with the left mouse will move the selection square to that position on the source. Dragging the mouse with the left button down will also move the selection square.

  Dragging the right mouse button will draw a focus square on the screen when you release the mouse your selection square will be resized to match the focus square. Optionally you can enter a new value in the text box in the grid tools. And then select the "Draw New Grid" command.

If you use the arrow buttons with "Enable Select Tool" unchecked it will move the selection square by one in that direction if you hold shift while pushing the arrow key the selection square will move based on its' size


------------------------------------------
other options and fixes include :
	Clicking or dragging over the ruler will color or select the row or column corresponding to that point.

	The undo feature has been updated to undo any edit. (NOTE: I have had problems with it but haven't
	Pinpointed the problem yet).

	I have fixed a problem with the resizing code that caused the app to crash if your screen wasn't a widescreen and you maxed the window.

	Fixed a problem when opening files that caused the program to crash.

	Fixed an error in the undo function that caused the rulers to flash.

	You can now unselect pixels by holding down the shift key while selecting.

	Error handlers have been inserted

	Replaced the commonDialog control with a class.

	Saving settings to an initial file.(NOTE: This can be changed to saving to the registry by simply removing the "GetSaveSettings" module). No other edits are necessary.

'after 1-4-08
	Fixed the pattern brush so when you select another color by selecting the pallet or by other means the brush is cleared.

	Added the ability to change the highlight (Selection) color

	Fixed an error saving and opening brush files.

	Added a check before an undo to see if we are in selection mode and a warning to indicate any selections will be cleared if the undo continues, with an option to abort.

	Fixed the Paste Rotate Function.

	Added a dialog to the rotate paste function allowing the user to enter the rotation amount and select the point in the selected image to rotate around. Also add a forward or reverse option to the rotation.

	Added a spin option to the rotation function that when chosen will allow you to enter an amount by which to increment the paste and an amount to stop the loop at and it will draw the image to be pasted at each interval.
	
	Fixed an error in the undo function that would overflow the integer used to count the changes. I changed it to double; you can change it to whatever.
___________________________________________________________________________________________________________________________
Known Issues
	There is a major drawback with the functions presented here in this app in that most options dealing with the pixels will ignore any black(0) pixels.

	The show color dialog isn't saving and loading the custom colors right
