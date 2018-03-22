A wrapper for Python's XlxsWriter

http://xlsxwriter.readthedocs.io/

It is approximately 4x slower than running the Python version

Installation
============

Step 1 - don't use windows

Step -1 oh no, I'm stuck on Windows

Pkg.add("PyCall")

This will, if you're fortunate, configure your Julia installation to use some version of Python.

It will print an info message on completion.

Here's my latest in full (I'm making a portable Julia on a USB flash so it's not using a .julia):

INFO: PyCall is using E:\DOT.Julia\v0.6\Conda\deps\usr\python.exe (Python 2.7.14) at E:\DOT.Julia\v0.6\Conda\deps\usr\python.exe, libpython = E:\DOT.Julia\v0.6\Conda\deps\usr\python27
INFO: E:\DOT.Julia\v0.6\PyCall\deps\deps.jl has been updated
INFO: E:\DOT.Julia\v0.6\PyCall\deps\PYTHON has been updated

The xlsx writer documentation lists a few ways to get it on your system

http://xlsxwriter.readthedocs.io/getting_started.html#getting-started


I don't have Pip or easy_install available on my Conda version above, so I shall download the tarball


It's Unix centric so you'll have to work out how to get the files out of the tarball yourself - I suggest 7Zip

I haven't worked out how to make my wrapper a proper public Julia package so just download it and put it in your Module path
Failing that use 

push!(LOAD_PATH, ".")

at the top of your script and keep XlsxWriter.jl in the same folder as your project

