# Python-and-COM
 How to let pythom make and use COM objects using win32com
 Making use of a python com object.
 The following code was built in PythonWIn running as an <b>administrator</b>.


 SimpleCOMServer.py - A sample COM server - almost as small as they come! 
 We simply expose a single method in a Python COM object.


    class PythonUtilities: 

       _public_methods_ = [ 'theSplitString' ]
      	_reg_progid_ = "PythonDemos.Utilities2"
     	 #  NEVER copy the following ID 
     	 #  Use "print pythoncom.CreateGuid()" to make a new one.
    	  _reg_clsid_ = "{492F4BC4- !!Dont use this number create your own!! 4-A79EE7EFFE35}"
    
  	       def theSplitString(self, val, item=None):
          import string
   	      resu=val.split()
   	      return resu


Add code so that when this script is run by
Python.exe, it self-registers.

    if __name__=='__main__':
       	print ("Registering COM server...")
  	     import win32com.server.register
   	    win32com.server.register.UseCommandLine(PythonUtilities)


!!Important !!! 
The __reg_clsid__ number was created using

     Import pythoncom
     Print (pythoncom.CreateGuid)

{????????????-Use this number -??????????}

After that run the above module and you should get no errors and should see 
Registered: PythonDemos.Utilities

On the interactive window.
In excel create a macro and call it TestPython.
Past make sure the macro code looks like this:

    Sub TestPython()

      Set PythonUtils = CreateObject("PythonDemos.Utilities2")
      response = PythonUtils.theSplitString("well well yes!")
    
        For Each Item In response
          MsgBox Item
        Next

      End Sub

