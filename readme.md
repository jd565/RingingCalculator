# Ringing Calculator

This is a program for receiving signals from bells and recording them. It then calculates statistics on the ringing.

The current code has:
An initial form which allows the user to input how many bells and COM ports are being used.
A form for setting up the COM ports by letting the user input the ports from a list.
Function for opening the ports.
A form for bell configuration
Functions to assign ports and pins to bells after the user has pressed a button.
A switch class to act as the start/stop switch
A bell class to hold everything to do with one bell.

Next steps:
Fill out the bell configuration form with extra settings
	debounce time
	handstroke and backstroke delay
Calculate statistics while the program is running
Function to print specific rows in the order they happened
save and load configuration