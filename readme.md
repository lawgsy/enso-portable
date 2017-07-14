# Description

A fork of [GChristensen]'s portable distribution of the community version of Humanized Enso Launcher for Windows, combined with [thdoan]'s branch.

The goal is to rewrite the code (to adhere to Python conventions and such), and get rid of whatever issues may be encountered.

# Installation instructions

Clone the repository. If using your own Python distribution rather than the one included, please make sure the code is written for Python 2.7.

# Running instructions

Run `debug.bat` (should be located in `enso/`), or compile `run-enso.exe` using the Visual Studio solution in `launcher/`.

Both require python (2.7) to be installed in a subdirectory `enso/python/` (this is to make the entire thing portable, without relying on Python begin installed).

Already have Python 2.7 as default python? `debug-globalpython.bat` instead (or modify the file to match your global installation).

# Known bugs

- Quasimode does not disappear upon loss of focus.

- Sometimes an Enso command will spawn another instance of Python.exe which is not closed later on. This instance will start to consume ridiculous amounts of CPU, as if it is in an event loop. It appears to happen when the window focus is changed while in quasimode. This also appears to happen mainly when 'go.py' commands are used, possibly because the enso window loses focus by definition there.

- The Enso tray icon does disappear upon exit due to a non-graceful exit.

- Capslock should be turned OFF upon exit.

# ToDo

- Add **.ensorc** configuration file (with settings such as default primary and secondary translation languages).

- Fix bugs.

# Copyright and licensing

See **LICENSE** file and credits/acknowledgements.

# Contact

E-mail: _lawgsy@gmail.com_

# Credits and acknowledgements

_Humanized, Inc._, _GChristensen_, _thdoan_ and all other people of the community that have in some way contributed to this project.

[gchristensen]: https://github.com/GChristensen/enso-portablet
[thdoan]: https://github.com/thdoan/enso-portable
