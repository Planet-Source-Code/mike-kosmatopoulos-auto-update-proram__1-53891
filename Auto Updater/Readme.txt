====Auto Update Program====

Didn't you always wanted to have an auto update form for your program?
Well now you only have to insert frmUpdate to your project, edit to suit your need and you are done!



Time to make it work 5-10 minutes

You need a server that supports HTML (like freepgs - www.freepgs.com)Usually every server supports HTML...

If you never had a free server then go and register a freepgs one.
Done registring? OK

1)Login to your freepgs account

2)Click "File Manager"

3)On the file manager down-left click "Create a File"

4)Put on Name of File update.html and on Title Auto updater

5)Click "Finish"

6)Back on file manager, click our new created file, update.html

7)Click Text Mode

8)Copy and Paste this: 

<HTML><HEAD><TITLE>Auto Updater</TITLE>

<META content="MSHTML 6.00.2800.1400" name=GENERATOR>
</HEAD>

<BODY>
Newest Version: <B>1.1</B>
<BR>
<BR>
What's New?: <B>Nothing is new since this is the first release of this program...%nOk?</B> 
<BR>
<BR>
Beta version?: <B>No</B> 
</BODY>
</HTML>



Don't forget to edit the values that are in <B> (bold) !!
Also if you want to have a vbCrLf (new line) in what's new, put %n and the program will automaticly replace it with vbCrLf .

9)Click Finish

10)Back at File Manager, click "Upload File(s)"

11)Click the first "Browse" button

12)Navigate the folder where the updated version of your program is

13)Click OK

14)Click "Upload"

15)Back at file manager, there should be a small icon near update.html .Click it.

16)On the new window, copy the text in the address bar

17)On VB at the AutoUpdate project look at the top. Edit "SiteInfo" to the text you copied.

18)Do the same for the updated version of your program.But instead of editing "SiteInfo" edit "DownloadSite"

19)Change CurrentVersion (again at the top of the source) to your program's curent version.

20)To make sure it works, change CurrentVersion to something else from the really current version.Press F5 and press "Check 
for Updates".It should say that there is a new version.When it asks you to update, click yes to make sure it works.If it doesn't, then try again...

If that is too complicated for you, please post the number of the step you have the problem and I'll try to explain better to you!


Also it is recommended that you attach MSINET.oxc when you release your program because many users don't have it...
The file is located at "x:\Windows\system32\" where x is the Hard Drive you are using. Also if you changed your windows directory durning windows setup you will have to replace it with the one you provided at the setup.



Please post comments!

Have Fun ;) 