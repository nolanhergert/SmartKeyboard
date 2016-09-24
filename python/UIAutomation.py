'''
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from PyQt4.QtWebKit import *
'''
import sys ,comtypes
from comtypes import *
from comtypes.client import *

# Windows dll's at https://en.wikipedia.org/wiki/Windows_API

# Get the graphics library. But none of them work!!
#comtypes.client.GetModule('gdi32.dll')
#comtypes.client.GetModule('comctl32.dll')
#from comtypes.gen.CommonControl import *

# Generate and import a wrapper module for the UI Automation type library,
# using UIAutomationCore.dll
comtypes.client.GetModule('UIAutomationCore.dll')
from comtypes.gen.UIAutomationClient import *



# Instantiate the CUIAutomation object and get the global IUIAutomation pointer:
iuia = CoCreateInstance(CUIAutomation._reg_clsid_,
                        interface=IUIAutomation,
                        clsctx=CLSCTX_INPROC_SERVER)

# raw = all the elements
rawView=iuia.rawViewWalker
# control = less of the elements
controlView=iuia.controlViewWalker



def ParseTreeDFS(tree):
    '''
    Walk the tree in depth-first search using recursion. Hopefully we don't 
    blow the stack! 
    '''
    # Get last element in tree and try and find children
    leaf = tree[-1]
    try:
        uiElement=controlView.GetFirstChildElement(leaf)
    except Exception as e:
        print e
        return
    while(True):
        try:
            # Do a quick test to check if it's not a null element
            temp = uiElement.CurrentName
            
            

                
            # Check if it's something we *can* click. By using the controlViewWalker
            # we are only getting things that are "controllable"/clickable??
            if uiElement.CurrentIsEnabled == 1:
            
                path = ''
                for element in tree:
                    if element.CurrentName == None:
                        path += '->'
                    else:
                        path += '->' + element.CurrentName.encode(sys.stdout.encoding,errors='replace')
                print path
            tree.append(uiElement)
            # Recurse! 
            ParseTreeDFS(tree)
            # Remove the item from the tree
            tree.pop()
            
        except ValueError, e:
            # Check if there's an error other than null pointer
            if str(e) != "NULL COM pointer access":
                # If so, continue raising the exception
                raise
            # Otherwise there's no valid element. We don't need to look for
            # more siblings so return
            break
        uiElement = controlView.GetNextSiblingElement(uiElement)
    return

    
def PrintPathToElement(uiElement):
    path = []
    
    while(True):
        path.append(uiElement)
        uiElement=controlView.GetParentElement(uiElement)
        # Make sure the uiElement isn't a null pointer
        try:
            temp = uiElement.CurrentName
        except ValueError, e:
            # Check if there's an error other than null pointer
            if str(e) != "NULL COM pointer access":
                raise
            break
    
    # Not sure how to make the path unique/traversable/reproducible among empty-string fields
    # and ID's that probably change with each program restart
    xpath = ''
    while(len(path) > 1):
        uiElement = path.pop()
        xpath += uiElement.CurrentName.encode(sys.stdout.encoding,errors='replace') + '->'
    xpath += path[0].CurrentName.encode(sys.stdout.encoding,errors='replace')
    print xpath
    
    
desktop = iuia.getRootElement()
ParseTreeDFS([desktop])

# Get the element that's underneath the mouse. It'd be nice to check this
# on a mouse event so that you can log what the user did
focused = iuia.getFocusedElement()
PrintPathToElement(focused)
print focused.currentName

#Doing help(comtypes.gen.UIAutomationClient) lets you know what is possible
# print help(comtypes.gen.UIAutomationClient)

'''
# Callbacks look really annoying to do in Python. Switch to C# at this point?
# The below doesn't work yet...
temp = comtypes.client.CreateObject(FocusChangedEventHandler)
iuia.AddFocusChangedEventHandler(temp)



 
class FocusChangedEventHandler(IUIAutomationFocusChangedEventHandler):

    def HandleFocusChangedEvent(element):
        print element.currentName

'''


