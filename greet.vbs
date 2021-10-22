Set Sapi = Wscript.CreateObject("SAPI.SpVoice")
 dim str
 if hour(time) < 12 then
 Sapi.speak "Good Morning Siddhant"         ' Add Your Own Greetings Message.
 else
 if hour(time) > 12 then
 if hour(time) > 16 then
 Sapi.speak "Good evening Siddhant"		   ' Add Your Own Greetings Message.
 else
 Sapi.speak "Good afternoon Siddhant"	   ' Add Your Own Greetings Message.
 end if
 end if
 end if
 Sapi.speak "The current time is"           ' Add Your Own Time Greetings Message.
if hour(time) > 12 then
 Sapi.speak hour(time)-12
 else
 if hour(time) = 0 then
 Sapi.speak "12"
 else
 Sapi.speak hour(time)
 end if
 end if
if minute(time) < 10 then
 Sapi.speak "o"
 if minute(time) < 1 then
 Sapi.speak "clock"
 else
 Sapi.speak minute(time)
 end if
 else
 Sapi.speak minute(time)
 end if
if hour(time) > 12 then
 Sapi.speak "P.M."
 else
 if hour(time) = 0 then
 if minute(time) = 0 then
 Sapi.speak "Midnight"
 else
 Sapi.speak "A.M."
 end if
 else
 if hour(time) = 12 then
 if minute(time) = 0 then
 Sapi.speak "Noon"
 else
 Sapi.speak "P.M."
 end if
 else
 Sapi.speak "A.M."
 end if
 end if
 end if
 
 
Dim max,min,choice

Dim arr

'Add a quote to the end of the this array 

arr= Array("Computers are useless.  They can only give you answers. Pablo Picaso","Never trust a computer you can’t throw out a window. Steve Wozniak","Controlling complexity is the essence of computer programming. Brian Kernigan","There are only two industries that refer to their customers as ‘users’. Edward Tufte","Most of you are familiar with the virtues of a programmer.  There are three, of course: laziness, impatience, and hubris. Larry Wall","Programmers are in a race with the Universe to create bigger and better idiot-proof programs, while the Universe is trying to create bigger and better idiots.  So far the Universe is winning. Rich Cook","The best programmers are not marginally better than merely good ones.  They are an order-of-magnitude better, measured by whatever standard: conceptual creativity, speed, ingenuity of design, or problem-solving ability. Randall E. Stross","There are only two kinds of programming languages: those people always bitch about and those nobody uses. Bjarne Stroustrup","There is no programming language–no matter how structured–that will prevent programmers from making bad programs. Larry Flon","Writing in C or C++ is like running a chain saw with all the safety guards removed. Bob Gray","In C++ it’s harder to shoot yourself in the foot, but when you do, you blow off your whole leg. Bjarne Stroustrup","Fine, Java MIGHT be a good example of what a programming language should be like.  But Java applications are good examples of what applications SHOULDN’T be like. pixadel","Good code is its own best documentation. Steve McConnell","Any code of your own that you haven’t looked at for six or more months might as well have been written by someone else. Eagleson’s Law","I don’t care if it works on your machine!  We are not shipping your machine! Vidiu Platon","Always code as if the guy who ends up maintaining your code will be a violent psychopath who knows where you live. Martin Golding","If debugging is the process of removing bugs, then programming must be the process of putting them in. Edsger W. Dijkstra","Avoiding complexity reduces bugs. Linus Torvalds")
min=0
max= UBound(arr) 'setting max as the size of array
Randomize
choice=Int((max-min+1)*Rnd+min) ' Using random funtion to select a quote index
sapi.speak arr(choice)


 
 
 
 
 