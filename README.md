wget.vbs
========

An HTTP file downloader (inspired by GNU [wget][]) written in VBScript.

[wget]: http://gnuwin32.sourceforge.net/packages/wget.htm


Rationale
---------

Ever needed to do some automated post-install and set-up on a brand-spankin'
new installation of Windows? Or, ever wished you could download stuff onto said
brand-new Windows box from the comfort and convenience of a batch script?
Yeah...

Hence wget.vbs, which is runnable on out-of-the-box set-ups. (It's in VBScript
instead of PowerShell because it was written in a day and an age when I was
still using Windows XP, but it will also work for Vista+)

You could even use VBS wget to get GNU [wget][]...

Use it wisely!


Requirements
------------

* cscript.exe (CLI)
* wscript.exe (GUI; default for double-click of VBS file)

...Both of which come with Windows, so nothing is really required (aside from
Windows itself) for wget.vbs to run. :-)


Usage
-----

CLI Execution:

    cscript wget.vbs <url> [save_to_file]

GUI Execution:

    wscript wget.vbs <url> [save_to_file]


License
-------

Boost Software License, Version 1.0: <http://www.boost.org/LICENSE_1_0.txt>


Acknowledgements
----------------

Because writing VBScript is hard, (mind-numbingly so,) I'd like to give a
grateful shout-out to these like-minded fellows:

* Vittorio Pavesi ([wget.vbs][1], Sep 2005)
* Chrissy LeMaire ([fileFetch.vbs][2], Jan 2007)
* James "HM2K" Wade ([wget.vbs][3], Dec 2009)
* Mahmoud Abu-Ghali ([wget.vbs][4], Apr 2011)
* Jean-Hervé Lescop and all the folks at [Adersoft][] for [VbsEdit][] (2001)

[1]: http://vittoriop77.altervista.org/vbscripts/wget.html
[2]: http://blog.netnerds.net/2007/01/vbscript-download-and-save-a-binary-file/
[3]: https://code.google.com/p/hm2k/source/browse/trunk/code/vbs/wget.vbs
[4]: http://abu-ghali.com/2012/04/11/wget-for-windows/
[Adersoft]: http://adersoft.com/
[VbsEdit]: http://vbsedit.com/
