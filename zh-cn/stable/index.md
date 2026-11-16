# xlwings - Make Excel Fly!

xlwings (Open Source) is a [BSD-licensed](http://opensource.org/licenses/BSD-3-Clause) Python library that makes it easy to call Python from Excel and vice versa:

* **Scripting**: Automate/interact with Excel from Python using a syntax close to VBA.
* **Macros**: Replace VBA macros with clean and powerful Python code.
* **UDFs**: Write User Defined Functions (UDFs) in Python (Windows only).

**Numpy arrays** and **Pandas Series/DataFrames** are fully supported. xlwings-powered workbooks are easy to distribute and work
on **Windows** and **Mac**.

Start here if you are new to xlwings. Learn about the syntax, the `RunPython` call, the add-in and UDFs.

More in-depths explanations about converters, debugging or how to write your own add-in.

To run xlwings without having to install Python, you have two options:

* [xlwings Lite](https://lite.xlwings.org): Install the free add-in from Excel's add-in store and you're done. It comes with a VS Code-like editor and Wingman, a powerful AI assistant.
* [xlwings Server](https://server.xlwings.org): Create your own modern Excel add-in with Python instead of JavaScript. Python runs on your own server and comes with enterprise features such as SSO and RBAC.

This is a description of all the classes, methods, properties and functions that xlwings offers to work with the Excel object model.
