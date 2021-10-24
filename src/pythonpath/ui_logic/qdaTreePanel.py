# -*- coding: utf-8 -*-
#!/usr/bin/env python

# =============================================================================
#
# Write your code here
#
# =============================================================================

import uno
import re
from functools import reduce

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import Rectangle
from com.sun.star.awt.tree import XTreeEditListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT
from com.sun.star.awt.MouseButton import RIGHT as MB_RIGHT

from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XSelectionSupplier
from com.sun.star.view.SelectionType import SINGLE as SELECTION_SINGLE

from ui.qdaTreePanel_UI import qdaTreePanel_UI

# ----------------- helpers for API_inspector tools -----------------

# uncomment for MRI
#def mri(ctx, target):
#    mri = ctx.ServiceManager.createInstanceWithContext("mytools.Mri", ctx)
#    mri.inspect(target)

# uncomment for Xray
#from com.sun.star.uno import RuntimeException as _rtex
#def xray(myObject):
#    try:
#        sm = uno.getComponentContext().ServiceManager
#        mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", uno.getComponentContext())
#        scriptPro = mspf.createScriptProvider("")
#        xScript = scriptPro.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
#        xScript.invoke((myObject,), (), ())
#        return
#    except:
#        raise _rtex("\nBasic library Xray is not installed", uno.getComponentContext())


class qdaTreePanel(qdaTreePanel_UI,XActionListener, XSelectionChangeListener, XTreeEditListener, XMouseListener):
    '''
    Class documentation...
    '''
    def __init__(self, panelWin):
        qdaTreePanel_UI.__init__(self, panelWin)

        # custom initialization
        self.TreeControl1.Editable = False  # True
        self.TreeControl1.InvokesStopNodeEditing = False

        self._contextMenu = None
        self._contextMenuContainer = {}
        self._contextMenuItems = {}

        self._tagIdents = {}  # maps tags to unique IDs
        self._lastTagID = 0

        # document
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        self.document = self.desktop.getCurrentComponent()
        self.updateTree() # this, seemingly, does not work

        # The dialog is created through models. To get the views/controllers we
        # need to call getControl(). To add to the irritation, there is the tree
        # model (model of the control element) and the tree data model (model of
        # what the control element shows).
        treeControl = self.DialogContainer.getControl('TreeControl1')
        treeControl.addSelectionChangeListener(self)
        treeControl.addTreeEditListener(self)
        treeControl.addMouseListener(self)

    def getHeight(self):
        return self.DialogContainer.Size.Height

    # --------- my code ---------------------
    # mri(self.LocalContext, self.DialogContainer)
    # xray(self.DialogContainer)

    def updateTree(self): #why is self implicitly passed?
        def convertAbstractToUiTree(abstractTree,parent,gui_treemodel):
            
            if not abstractTree: #if abstract tree is empty, show some into to the user
                branch = treemodel.createNode("if you create comments write a #hashtag, they will be listed here", False)
                parent.appendChild(branch)
                
            for item in abstractTree: #TODO: Item is sometimes the string "children"
                
                if "children" in item and item["children"]:
                    branch = treemodel.createNode("#"+item['name'], True) #[:n] takes the first n chars of a string (or just leaves it be, if there are less). Alternatives: if-clause or function for it; gain: add "…" to string if shortened 
                else:
                    if len(item['name']) > 20:
                        branch = treemodel.createNode(item['name'][:20]+'...', False)
                    else:
                        branch = treemodel.createNode(item['name'], False)
               
                if "children" in item and item["children"]: #"in" checks for keys existance, "item["children"]"" checks if list is emptry since an empty list is false (https://www.python.org/dev/peps/pep-0008/#programming-recommendations)
                    convertAbstractToUiTree(item["children"],branch, gui_treemodel)
                    
                if "data" in item:
                    branch.DataValue = item["data"]["origAnnotation"]
                parent.appendChild(branch)

        treeControl = self.DialogContainer.getControl('TreeControl1')
        commentslist = self._collectHashtaggedComments()
        abstractTree = constructTree(commentslist)
        sortTreeRecursive(abstractTree) #inplace sort

        treemodel = self.ServiceManager.createInstance("com.sun.star.awt.tree.MutableTreeDataModel")

        rootnode = treemodel.createNode("root",True)
        treemodel.setRoot(rootnode)

        convertAbstractToUiTree(abstractTree, rootnode, treemodel)

        self.TreeControl1.DataModel = treemodel

        toolkit = self.ServiceManager.createInstance("com.sun.star.awt.Toolkit")

        expandAllNodesGuiTree(treeControl.Model.DataModel.Root, treeControl)

        self.TreeControl1.SelectionType = SELECTION_SINGLE
        self.TreeControl1.RootDisplayed = False

        expandAllNodesGuiTree(treeControl.Model.DataModel.Root, treeControl)

    def _collectHashtaggedComments(self):
        # DOES: Collect all comments ("Annotations") that match a regex in an array
        # RETURNs: List of comments ("Annotations")

        # This method collects all comments and inserts a list of unique tag IDs
        # into the author field. This way, the annotations get colored by tags.
        # Annoyingly, this requires saving and reloading the document because
        # the "post-its" don't get updated when the author changes.
        #
        # I tried replacing annotations with rewritten versions but that does not
        # work. 1) It is impossible to retain the selection when calling
        # document.Text.insertTextContent(anchor, newField, True). If the last
        # parameter is "True", LO becomes highly unstable.
        # 2) Comment threads would be lost when re-adding comments.

        document = self.document

        # This regex needs to match for a comment to be included in the returned list.
        findTagsRe = re.compile(r'#\S+')
        authorRe = re.compile(r' {([0-9]+)(\+[0-9]+)*}$')

        textFields = document.getTextFields()
        matchedComments = []

        for count, currentField in enumerate(textFields):
            if not currentField.supportsService("com.sun.star.text.TextField.Annotation"):
                continue  # field is not a comment

            if not findTagsRe.search(currentField.Content.strip()):
                continue  # field contains no tags

            allTags = sorted([str(x).lower() for x in findTagsRe.findall(currentField.Content.strip())])  # e.g. ['#tag1#nested1#nested2', '#tag2']

            for tag in allTags:
                if tag not in self._tagIdents:
                    self._lastTagID += 1
                    self._tagIdents[tag] = self._lastTagID

            splitTags = [x[1:].split("#") for x in allTags]  # e.g. [ ['tag1', 'nested1', 'nested2'], ['tag2'], ]
            markedText = currentField.getAnchor().getString()
            taggedAuthor = ' {'+("+".join([str(self._tagIdents[x]) for x in allTags]))+'}'

            if authorRe.search(currentField.Author):
                currentField.Author = authorRe.sub(taggedAuthor, currentField.Author)
            else:
                currentField.Author += taggedAuthor

            annotationInfo = {
                'name': markedText,
                'paths': splitTags,
                'data': {
                    'id': count,
                    'origAnnotation': currentField,
                    'markedText': markedText,
                    'content': currentField.Content,
                    }
                }

            matchedComments.append(annotationInfo)

            print("collected:", markedText)

        return matchedComments

    # --------- helpers ---------------------

    def messageBox(self, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
        sm = self.LocalContext.ServiceManager
        si = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.LocalContext)
        mBox = si.createMessageBox(self.Toolkit, MsgType, MsgButtons, MsgTitle, MsgText)
        return mBox.execute()

    def _showContextMenu(self, node):
        self._contextMenuItems = {  # inefficient but easier to change
                'dataNode': ['Move', 'Delete'],
                'tagNode': ['Edit', 'Delete'],
            }

        if node.DataValue:
            kind = 'dataNode'
        else:
            kind = 'tagNode'

        self._createContextMenu(kind)
        if not self._contextMenu:
            return

        comp = self.document.getCurrentController().getFrame().getComponentWindow()
        # position of the popupmenu is not considered
        if n := self._contextMenu.execute(comp, Rectangle(), 0):
            print(f"- selected: {n} -> {self._contextMenuItems[kind][n-1]}")

            if kind == 'dataNode':
                if n == 1:
                    pass
                elif n == 2:
                    ret = self.messageBox(f'Are you sure you want to delete the tagging of "{node.getDisplayValue()}"?', 'Confirm', WARNINGBOX, BUTTONS_OK_CANCEL)

                    if ret == 1:  # accepted
                        pass
            elif kind == 'tagNode':
                if n == 1:
                    pass
                elif n == 2:
                    ret = self.messageBox(f'Are you sure you want to delete the tag "{node.getDisplayValue()}" '+
                                           'and all information associated with it?', 'Confirm',
                                           WARNINGBOX, BUTTONS_OK_CANCEL)

    def _createContextMenu(self, kind):
        if kind in self._contextMenuContainer:
            self._contextMenu = self._contextMenuContainer[kind]
            return

        smgr = self.ctx.getServiceManager()
        popup = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", self.ctx)

        if kind not in self._contextMenuItems:
            self.messageBox(f"Bug: unknown context menu '{kind}'", "Error", ERRORBOX)
            return

        for i, item in enumerate(self._contextMenuItems[kind]):
            print(f"adding: {i+1} / {item} / {i}")
            popup.insertItem(i+1, item, 0, i+1)

        self._contextMenuContainer[kind] = popup
        self._contextMenu = popup

    # -----------------------------------------------------------
    #               Execute dialog
    # -----------------------------------------------------------

    def showDialog(self):
        self.DialogContainer.setVisible(True)
        self.DialogContainer.createPeer(self.Toolkit, None)
        self.updateTree() # can now execute the update (and within expand the nodes) since the peer is set now, and this is needed for expanding nodes (whyever...)
        self.DialogContainer.execute()

    # -----------------------------------------------------------
    #               Action events
    # -----------------------------------------------------------


    def actionPerformed(self, oActionEvent):
        if oActionEvent.ActionCommand == 'updateButton_OnClick':
            self.updateButton_OnClick()

    def updateButton_OnClick(self):
        self.updateTree()
#         self.DialogModel.Title = "It's Alive! - updateButton"
#         self.messageBox("It's Alive! - updateButton", "Event: OnClick", INFOBOX)

    def selectionChanged(self, ev):
        selection = ev.Source.getSelection().DataValue #get id of item

        if selection is None:
            return  # hashtag node

        self.selection = selection
        self.textOfTag.Text = selection.getAnchor().getString()
        #mri(self.LocalContext, XSCRIPTCONTEXT.getDocument())

        if self.CheckboxJumpto.State == True:
            scrollToRange(self.document, selection.getAnchor())

    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/tree/XTreeEditListener.html
    def nodeEditing(self, node):
        # self.TreeControl1.cancelEditing()
        # self.messageBox("asdölkjasdölkj: "+node.getDisplayValue(), "title")
        pass

    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/tree/XTreeEditListener.html
    def nodeEdited(self, node, newText):
        if node.getDisplayValue() != newText:
            self.messageBox("done: "+newText, "title")

    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/XMouseListener.html
    def mousePressed(self, ev):
        return False

    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/awt/XMouseListener.html
    def mouseReleased(self, ev):
        if ev.Buttons == MB_RIGHT:
            print("requesting context menu")

            try:
                node = self.DialogContainer.getControl('TreeControl1').getNodeForLocation(ev.X, ev.Y)
            except:
                print(f"failed to select node at {ev.X}x{ev.Y}")

            if node:
                print("node found at", ev.X, ev.Y)

                treeControl = self.DialogContainer.getControl('TreeControl1')
                treeControl.select(node)

                # self.messageBox(f"left double click at {ev.X}x{ev.Y}: {node.getDisplayValue()}", "edit now")
                self._showContextMenu(node)
            else:
                print(f"no node at {ev.X}x{ev.Y}")

        return False

#--------------------------------------------

def scrollToRange(document, range): #range is a section in a text
    #DOES: scrolls the document to the range in a given text
    #ARGUMENTS: a document and a range object. The range must be in the document
    #RETURNS: nothing (is a side effect function, see DOES)

    viewCursor = document.CurrentController.getViewCursor()
    viewCursor.gotoRange(range, False)

    # It is not necessary to collapse (clear) the selection, as accidental
    # keystrokes only reach the currently focussed item.
    # viewCursor.collapseToEnd()


def constructTree(commentsArray):
    '''
    DOES: Create a nested tree from a flat list 
    GETs: A list with items having name (string), paths (list of paths each being a list composed of path parts), data (a dict with what you like – it’s payload)
    RETURNS: a tree like this
    
    - A
     * Children:
      - B
       * children:
        - D
        - E
      - C
    …
    So it is lists of dicts. One property is "children", in which there is another list of dicts etc.
    '''
    
    #TODO: Fix "array" name to "list"  
    
    result = []

    initialValue = {
        0:result
    }

    for comment in commentsArray:
        for path in comment["paths"]:
            #pydevd.settrace()
            def reducerfunction(accumulator,pathpart): #function defined here so I can access the current comment in the reducer function
                #print("cmmt", comment, "-- -- - -path",path, "- - - pathpart", pathpart)
                if not pathpart in accumulator:
                    accumulator[pathpart] = {
                        0:[]
                    }

                    elementToAppend = {
                        "name":pathpart,
                        "children":accumulator[pathpart][0]
                    }

                    if pathpart == comment["name"] and "data" in comment:
                        elementToAppend["data"] = comment["data"]

                    accumulator[0].append(elementToAppend)
                return accumulator[pathpart]

            reduce(reducerfunction,path+[comment["name"]],initialValue)

    return result


def sortTreeRecursive(treeList):
    ''' 
    DOES: Sort children lists in a tree by name
    GET: Tree
    RETURN: Nothing, side effect
    '''
    for item in treeList:
        if "children" in item:
            sortTreeRecursive(item["children"])
    treeList.sort(key=lambda item:item["name"])


def expandAllNodesGuiTree(root,treeControl):
    '''
    DOES: Expand all Nodes in a mutableTreeModel
    GETS: XTreeNode
    RETURNS: Nothing, side effect
    '''

    # TODO: check if tree control has a peer, if not return None
    for count in range(0,root.ChildCount):
        child = root.getChildAt(count)
        
        if child.ChildCount > 0 and treeControl.Peer:
            treeControl.expandNode(child)
            expandAllNodesGuiTree(child, treeControl)

        
    
    


#----------------#
#-- RUN PANEL----#
#----------------#

def Run_qdaTreePanel(*args):
    """
    Intended to be used in a development environment only
    Copy this file in src dir and run with (Tools - Macros - MyMacros)
    After development copy this file back
    """
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    dialog = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)

    app = qdaTreePanel(dialog)
    app.showDialog()

g_exportedScripts = Run_qdaTreePanel,
