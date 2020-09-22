ActiveX Exe MultiThreading Example - By Lewis Miller

-Introduction-
Alot of vb programmers have little idea of the power and capabilities of activex exe projects and how they work, or even how threading works or can be accomplished. This project attempts to show how easy it is to actually use an activex exe to multithread. Many have probably seen the examples available that use complicated api calls and lost of classes, and are prone to crashing or are limited in functionality. Not so with an activeX exe!! You can fully test and debug your active exe's from a separate instance of VB! Simply start the ActiveX Exe with the 'Start With Full Compile' option (Ctrl+F5) and then run your main exe project that references the active x.

-Some Notes-
Many c/c++ programmers scoff at vb because they say it cant multithread and is therefore useless for large scale applications. Now you can see there comments and silently giggle as you think about them spending weeks synchronizing and coding there 'threaded' c++ application with lots of api calls, cryptic mutex's and semaphores. Because you know you can easily create a separate thread literaly within minutes/hours (with this template i can usually do it in less than an hour depending on the work to be done). We VB programmers can let microsoft invent the wheel once and re-use it through the power of COM (Component Object Model) and ActiveX Exe's.

Keep in mind though that multithreading does not nessacarily fix bad coding, and isnt the end all answer to all bottlenecks in code. Poor coding is simply that, poor coding, and 20 threads wont help it be any faster. Many times new threads arent needed to complete a long task if its properly coded. I personally have coded a chatserver that is NOT multithreaded, in visual basic 6, and it services an average of 500 logins an hour, but only uses 20mb of ram and an avg of 10% cpu useage with 800 busy users logged. It is a good example of what good optimized code can do without many threads.
Here is a general guidline of when a new thread will benefit your application:
1) If a single task takes longer than 2 seconds then it could possibly benefit from another thread.
2) If the task doesnt require any input after it is working, (its a pain to add more input)
3) If your application has a gui that needs to stay responsive while processing a long task

The overhead required by COM to marshal data to a new thread isnt worth it if the task to be done doesnt take at least a second or more to complete. So you should weigh the benefits against the drawbacks. The neatest thing about a an active x exe is that you can actually put them on a seperate computer (and have several to create a server farm, not covered here though) to extend the power even more to multiple processors (through DCOM - Distributed Component Object Model).

-Testing The Project-
1) Unzip the source code to a folder, making sure to preserve directorys. There are 2 separate projects, and one is in a seperate folder.
2) Open the activex threader project 'prjThreader.vbp' in vb and press ctrl+F5 to run it (start with full compile)
3) Open the 'prjTest.vbp' project in vb and goto Project > References and make sure there is a reference to 'ActiveX Threader', if not make sure it is. You may have to uncheck it, save the project, close and re-open it, and recreate the reference again.
4) run the test project, and check out Example 2, to see how multi threading can work

enjoy :)
