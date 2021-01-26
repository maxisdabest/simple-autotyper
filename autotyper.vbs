set shell = createobject ("wscript.shell")

strtext = inputbox ("Write down your message you like to spam")
strtext2 = inputbox ("Message 2")
strtext3 = inputbox ("Message 3")
strtext4 = inputbox ("Message 4")
strtext5 = inputbox ("Message 5")
strtimes = inputbox ("How many times do you like to spam? Type in 1 less than the amount of msgs you want to spam. eg; if want to spam 1000, type 999.Max no.1999")
strspeed = inputbox ("Type 1000 for 1 msg per second 100 for 10 msg per second etc type 1 for fastest")
strtimeneed = inputbox ("How many SECONDS do you need to get to your victims input box? Click where you want to spam.")
If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "You entered something else then a number on Times, Speed and/or Time need. shutting down"
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "You have " & strtimeneed & " seconds to get to your input area where you are going to spam."
wscript.sleep strtimeneed2
for i=0 to strtimes
shell.sendkeys (strtext & "{enter}")
shell.sendkeys (strtext2 & "{enter}")
shell.sendkeys (strtext3 & "{enter}")
shell.sendkeys (strtext4 & "{enter}")
shell.sendkeys (strtext5 & "{enter}")
wscript.sleep strspeed
Next
wscript.sleep strspeed * strtimes / 10

returnvalue=MsgBox ("Want to spam again?",36)
If returnvalue=6 Then
End If
If returnvalue=7 Then
msgbox "Spambox is shutting down"
wscript.quit
End IF
strtext = inputbox ("Write down your message you like to spam")
strtimes = inputbox ("How many times do you like to spam? Type in 1 less than the amount of msgs you want to spam. eg; if want to spam 1000, type 999")
strspeed = inputbox ("Type 1000 for 1 msg per second 100 for 10 msg per second etc type 1 for fastest")
strtimeneed = inputbox ("How many SECONDS do you need to get to your victims (tom's) input box? Click where you want to spam")
loop