# VBA-StateLossCallback
A class that allows safe callbacks when state is lost.

- Each instance of this class will make a call back to the provided macro when state is lost (Application exists, Stop button is pressed in IDE, Desgin Mode etc.)
- **No** memory leaks, **no** crashes
- Compatible with Windows and Mac on both x32 and x64

## Implementation
An extra interface is used to achieve safety. The ```IUnknown::Release``` method in the extra virtual table is replaced with one of the methods of the extra interface which then performs the call on the actual method implementation. This 'wrapping' is what makes the call safe even if the IDE Stop button is pressed while the callback is performed. There is a single instance of the extra interface (see ```m_data``` class member) and it is unmanaged i.e. no actual reference is added or removed from the object reference count.

Design decisions:
- to keep the solution compact, an extra custom interface was avoided. However, this required another interface that is already available and so the ```MSForms.DataObject``` interface was chosen as it's ```Clear``` method has no parameters and is compatible (same signature) with ```Release```. All the other interface methods are not used but had to be implemented so that the solution compiles. In other words, there are a few unused methods but the solution has one less class to distribute
- since the user can press Stop during the callback, the solution could not rely on restoring the virtual table functions and so the solution must work even after the table has been altered. Because of that, there is no call made to ```AddRef``` or ```Release``` and the ```m_data``` instance is unmanaged. One way of getting the extra interface pointer is to do something like ```Set m_data = Me``` followed by ```ObjPtr(m_data)``` but both these calls would make use of ```AddRef``` and ```Release```. Instead we get the interface pointer using ```VarPtr(DataObject) + PTR_SIZE``` which is perfectly valid - note that the variable sits on the line ```Implements MSForms.DataObject ``` which is a very useful trick
- the callback can be one of the two:
  - a macro name passed to ```InitByMacroName```. Up to 30 callback arguments are also allowed. Objects should be avoided as arguments as there is no guarantee they will still be 'alive' after state loss
  - an AddressOf pointer passed to ```InitByAddress```. One text callback argument allowed
- for the callback by address to work, a second extra interface is used (```stdole.IFontEventsDisp```) but this time the instance of this interface is actually managed and the ```FontChanged``` method is used as the safe wrapper to call the desired callback method

## Installation
Just import the following code modules in your VBA Project:
* [**StateLossCallback.cls**](https://github.com/cristianbuse/VBA-StateLossCallback/blob/master/src/StateLossCallback.cls)

## Demo
Import the following code module from the [demo folder](https://github.com/cristianbuse/VBA-StateLossCallback/tree/master/src/Demo) in your VBA Project:
* [Demo.bas](https://github.com/cristianbuse/VBA-StateLossCallback/blob/master/src/Demo/Demo.bas) - run the ```Demo#``` methods

## License
MIT License

Copyright (c) 2023 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
