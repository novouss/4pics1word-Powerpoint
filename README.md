# 4pics1word Powerpoint

4-pics-1-word is a guessing game with four images and the player has to guess a specific word related to the four pictures presented. 

The player is given an image, a keyboard containing a limited amount of letters, a label that shows their input, and a button that clears their input. The player is given an unknown word that can only be described with the four photos they're given. If the player inputs the incorrect answer say "WRONG!", else if they got the answer correctly say "CORRECT!". Upon getting the correct answer the player is given a NEXT button that moves them to the next level/slide. Everything in the previous slide is resetted.

# How to use

1. Upon openning the powerpoint file ensure that you've `Enabled Editing` on the yellow ribbon presented to you.
2. Go to `File > Options > Customize Ribbon > Check "Developer Mode" on Main Tabs.` This is for future edits and iterations of the program.
3. Go to `File > Options > Trust Center > Trust Center Settings > Click on "Enable all macros".` This is for buttons and text to function.
4. Restart Powerpoint by closing the window and openning the file.
5. Start presenting and everything should work as intended.

# Creating my own 4-pics-1-word game

1. On your Ribbon, open the newly added `Developer` tab and click on `Visual Basic.` A new window called `Microsoft Visual Basic for Applications` should open.

*Note: If you don't have the `Developer` tab go to the **How to use** section and follow "Step 2."*

2. On the `Project - VBAProject` window on the left side of your screen, click on `Slide2.`
3. Under the line `Private Function reset()` uncomment `ANSWER.Visible = True` by removing the single quotation symbol (') beside it. 
4. Save the changes by clicking on `Ctrl + S` and present the second slide in powerpoint.

*Note: Slides in VBA and Powerpoint are connected, Slide1 = the first slide, Slide2 = second slide, etc.*

6. A new `box` should appear, this is the answer box, and this will be the word players have to guess.
7. To edit, click on the `box` on powerpoint, on the `Developer` tab click on `Properties.` Find the `Caption` row, this is where you can edit the text.
8. After editing the text be sure to also edit the text on the buttons for users to get the correct answer.
9. To edit the text on the buttons, follow the same procedure on "Step 6". but click on the buttons instead of the box.
10. After making your edits, go back to the `Microsoft Visual Basic for Applications` window and comment `ANSWER.Visible = True` by putting a single quotation sumbol (') besite it like before. This is to hide the answer.
11. Save the changes by clicking on `Ctrl + S` and present the second slide in powerpoint. Test the buttons and try to get wrong answers and the correct answer.

*Note: Each slide contain the exact same lines of code. You'll need to individually uncomment `ANSWER.Visible = True` if you want to change the answer to that slide.*

# "There's not enough buttons for the word I want to use"

1. On Powerpoint click on the buttons and copy either an entire row or the amount you want to use.
2. (Optional) Click on one of the buttons, on the `Developer` tab click on `Properties.` Find the `(Name)` row, change the name to `btn11 or btn(n + 1) (let n be the number of buttons you currently have)`. 
3. Double click on the button and you should be brought back to `Microsoft Visual Basic for Applications.`
4. You'll be given a new line and function named `Private Function btn#_Click()` or `Private Function CommandButton#_Click()`
5. Under this Function simply type `userInput btn#` or `userInput CommandButton#`. 
6. Under the line `Private Function btn_enable()`, add `btn#.Enabled = True` or `CommandButton#.Enabled = True`. This is for the program to reset the buttons.

*Note: The pound sign (#) refers to its button number, it could be btn1, btn2, etc.. Just copy the button number on `btn#_Click()` and put the same number on `userInput btn#` and else where.

5. Save the changes by clicking on `Ctrl + S` and present the second slide in powerpoint. Test the buttons.
