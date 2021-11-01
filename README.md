# XLSXWriter - it's working again!

This is a wrapper around 

https://github.com/jmcnamara/XlsxWriter

which is the *best* Excel writing module for Python, hands down.

Today I finally got it to work as a proper Julia package so you can just Pkg.add("https://github.com/lawless-m/XlsxWriter.jl") and start using it, after a little more setup.

To make it work, it is up to you, dear user, to arrange for the Python code to be available

I've used an environment variable for the task : ENV["XLSXWRITER_PATH"]

and this should be set to the directory inside which the directory xlsxwriter can be found

So, for example, if you Clone the above code to /opt/XlsxWriter.py then set ENV["XLSXWRITER_PATH"] = "/opt/XlsxWriter.py"

otherwise Julia will throw an error to tell you to do that.

Excitingly, here in November 2021, Release 3.0.2 is coming up, with new stuff in it, so it looks like I have some work to do!



