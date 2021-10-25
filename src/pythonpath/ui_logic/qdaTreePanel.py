# -*- coding: utf-8 -*-
#!/usr/bin/env python

# =============================================================================
#
# Write your code here
#
# =============================================================================

import uno
import os
import re
import sys
import tempfile
import colorsys
from collections import defaultdict

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxButtons import DEFAULT_BUTTON_OK, DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX

from com.sun.star.document import XEventListener
from com.sun.star.beans import PropertyValue
from com.sun.star.beans.PropertyState import DIRECT_VALUE

from com.sun.star.awt import XActionListener
from com.sun.star.awt import XMouseListener
from com.sun.star.awt import Rectangle
from com.sun.star.awt.tree import XTreeEditListener
from com.sun.star.awt.MouseButton import LEFT as MB_LEFT
from com.sun.star.awt.MouseButton import RIGHT as MB_RIGHT
from com.sun.star.awt.FontWeight import BOLD as FW_BOLD

from com.sun.star.view import XSelectionChangeListener
from com.sun.star.view import XSelectionSupplier
from com.sun.star.view.SelectionType import SINGLE as SELECTION_SINGLE

from hsluv import hsluv
from ui.qdaTreePanel_UI import qdaTreePanel_UI


# GLOBAL STATIC CONFIG
PACKAGE_ID = 'tagtree.qdaaddon.de.fordes.qdatreehelper'


def get_traceback():
    """
    Get a traceback for pyuno exceptions.
    Source: https://forum.openoffice.org/en/forum/viewtopic.php?f=20&t=69813#p311800
    """
    (excType, excInstance, excTraceback) = sys.exc_info()
    ret = (
        str(excType) + ": " +
        str(excInstance) + "\n" +
        uno._uno_extract_printable_stacktrace(excTraceback)
    )
    return ret


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


class Tree(defaultdict):
    def __init__(self, parent):
        super().__init__(parent)
        self.children = []
        self.path = ''

    @property
    def name(self):
        return '#'+self.path.split('#')[-1]


def NestedTree():
    return Tree(NestedTree)


class Leaf():
    def __init__(self, path, data):
        self.path = path
        self.data = data

    @property
    def pathString(self):
        return '#'+('#'.join(self.path))

class Annotation():
    def __init__(self, ident, name, paths, textField):
        self.ident = ident
        self.name = name
        self.allPaths = paths
        self.textField = textField

    @property
    def content(self):
        return self.textField.Content

    @property
    def text(self):
        return self.textField.getAnchor().getString()


class qdaTreePanel(qdaTreePanel_UI,XActionListener, XSelectionChangeListener, XTreeEditListener, XMouseListener, XEventListener):
    '''
    Class documentation...
    '''
    FIND_TAGS_RE = re.compile(r'#\S+')
    AUTHOR_TAG_IDS_RE = re.compile(r' {([0-9]+)(\+[0-9]+)*}$')

    def __init__(self, panelWin):
        qdaTreePanel_UI.__init__(self, panelWin)

        # custom initialization
        self.TreeControl1.Editable = False  # True
        self.TreeControl1.InvokesStopNodeEditing = False
        self.TreeControl1.SelectionType = SELECTION_SINGLE
        self.TreeControl1.RootDisplayed = True

        self._objectsCache = {}
        self._contextMenu = None
        self._contextMenuContainer = {}
        self._contextMenuItems = {}

        # document pointers
        self.ctx = uno.getComponentContext()
        self.smgr = self.ctx.ServiceManager
        self.desktop = self.smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        self.document = self.desktop.getCurrentComponent()

        self.globalEvents = self.smgr.createInstanceWithContext("com.sun.star.frame.GlobalEventBroadcaster", self.ctx)
        self.globalEvents.addEventListener(self)

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

    def updateTree(self):
        # reset aggregated data
        self._tagIdents = {}
        self._lastTagID = 0
        self._objectsCache = {}

        # reset document pointers
        # This is required to support working on multiple documents at the same time.
        self.document = self.desktop.getCurrentComponent()

        treeControl = self.DialogContainer.getControl('TreeControl1')
        commentslist = self._collectTaggedComments()
        abstractTree = self._constructTree(commentslist)
        treemodel = self.ServiceManager.createInstance("com.sun.star.awt.tree.MutableTreeDataModel")
        rootnode = treemodel.createNode("QDA Tags", True)
        treemodel.setRoot(rootnode)

        def sortTreeRecursive(tree):
            # don't sort children -- tree.children.sort(key=lambda x: x.data.name)
            for key, value in sorted(tree.items(), key=lambda item: item[1].name):
                tree[key] = tree.pop(key)  # hackish way to sort our custom dict in-place
            for i in tree.items():
                sortTreeRecursive(i[1])
        sortTreeRecursive(abstractTree)  # TODO make sorting optional

        self._convertAbstractToUiTree(abstractTree, rootnode, treemodel)
        self.TreeControl1.DataModel = treemodel
        self._expandAllNodesGuiTree(treeControl.Model.DataModel.Root, treeControl)

    def _collectTaggedComments(self):
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

        textFields = document.getTextFields()
        matchedComments = []
        count = -1  # will be reset if the loop runs

        for count, currentField in enumerate(textFields):
            if not currentField.supportsService("com.sun.star.text.TextField.Annotation"):
                continue  # field is not a comment

            if not self.FIND_TAGS_RE.search(currentField.Content.strip()):
                continue  # field contains no tags

            # collect all fields that match a regex
            # result: e.g. ['#tag1#nested1#nested2', '#tag2']
            allTags = sorted([str(x).lower() for x in self.FIND_TAGS_RE.findall(currentField.Content.strip())])

            for tag in allTags:
                if tag not in self._tagIdents:
                    self._lastTagID += 1
                    self._tagIdents[tag] = self._lastTagID

            splitTags = [x[1:].split("#") for x in allTags]  # e.g. [ ['tag1', 'nested1', 'nested2'], ['tag2'], ]
            markedText = currentField.getAnchor().getString()
            taggedAuthor = ' {'+("+".join([str(self._tagIdents[x]) for x in allTags]))+'}'

            if currentAuthor := self.AUTHOR_TAG_IDS_RE.search(currentField.Author):
                if currentAuthor.group(0) != taggedAuthor:
                    currentField.Author = self.AUTHOR_TAG_IDS_RE.sub(taggedAuthor, currentField.Author)
            else:
                currentField.Author += taggedAuthor

            matchedComments.append(Annotation(
                ident=count,
                name=markedText,
                paths=splitTags,
                textField=currentField,
                ))

        print(f"collected {count+1} tagged comments")
        return matchedComments

    def _constructTree(self, commentsList):
        '''
        DOES: Create a nested tree from a flat list
        GETs: A list of Annotation objects
        RETURNS: a tree like this:

        - Tree A (tag):
            * path: ''
            * children: [Leaf(path, Annotation), Leaf(...)]
            * items: {Tree...}
        '''
        tree = NestedTree()

        for comment in commentsList:
            for path in comment.allPaths:
                subtree = tree
                pathShard = ''

                for part in path:
                    subtree = subtree[part]
                    pathShard += '#'+part
                    subtree.path = pathShard

                subtree.children.append(Leaf(path=path, data=comment))

        return tree

    def _convertAbstractToUiTree(self, abstractTree, parent, treemodel):
        if not abstractTree:
            # show documentation if there are no nodes
            branch = treemodel.createNode("Tagged comments will be listed here.", False)
            exampleA = treemodel.createNode("Tags may be simple: #tag", False)
            exampleB = treemodel.createNode("Or nested: #tag#subtag#subsubtag", False)
            exampleC = treemodel.createNode("Comments may contain multiple tags.", False)

            parent.appendChild(branch)
            branch.appendChild(exampleA)
            branch.appendChild(exampleB)
            branch.appendChild(exampleC)

        if not parent.DataValue:
            self._objectsCache[id(abstractTree)] = abstractTree
            parent.DataValue = id(abstractTree)

        for item in abstractTree.values():
            self._objectsCache[id(item)] = item
            branch = treemodel.createNode(item.name, True)
            branch.DataValue = id(item)

            for child in item.children:
                self._objectsCache[id(child)] = child
                leaf = treemodel.createNode(child.data.name[:28]+('...' if len(child.data.name) > 28 else ''), False)
                leaf.DataValue = id(child)
                branch.appendChild(leaf)

            if item:  # dict has nested items
                self._convertAbstractToUiTree(item, branch, treemodel)

            parent.appendChild(branch)

    def _expandAllNodesGuiTree(self, root, treeControl):
        '''
        DOES: Expand all Nodes in a mutableTreeModel
        GETS: XTreeNode
        RETURNS: Nothing, side effect
        '''
        treeControl.expandNode(root)

        for count in range(0, root.ChildCount):
            child = root.getChildAt(count)

            if child.ChildCount > 0 and treeControl.Peer:
                treeControl.expandNode(child)

            self._expandAllNodesGuiTree(child, treeControl)

    # --------- reports ---------------------

    def _buildTagDataList(self, abstractTree, tempList=None):
        if tempList is None:
            tempList = []

        if hasattr(abstractTree, 'children'):
            for child in abstractTree.children:
                tempList.append(child)

        if not abstractTree:
            return tempList

        for item in abstractTree.values():
            tempList = self._buildTagDataList(item, tempList)

        return tempList

    def _writeToTableCell(self, table, column, row, content):
        cell = table.getCellByPosition(column, row)
        cursor = cell.createTextCursor()
        cell.insertString(cursor, content, False)

    def _createTagReport(self, tag):
        print("building report for tag:", tag.name, '|', tag.path)

        packageInfo = self.ctx.getByName('/singletons/com.sun.star.deployment.PackageInformationProvider')
        packageDir = packageInfo.getPackageLocation(PACKAGE_ID) if packageInfo else None

        if not packageDir:
            print("error: failed to get package directory")
            self.messageBox('Internal error: failed to determine template directory.',
                            'Internal Error', ERRORBOX)
            return

        try:
            # blank document: report = self.desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
            loadProperties = [PropertyValue("AsTemplate", 0, True, DIRECT_VALUE)]
            report = self.desktop.loadComponentFromURL(packageDir+'/templates/tag_report.odt', "_blank", 0, tuple(loadProperties))
            table = report.getTextTables().getByName('TAG_TABLE')
        except:
            print("error: failed to load template:", get_traceback())
            return

        if not table:
            print("error: no table named TAG_TABLE found in template")
            return

        dataNodes = self._buildTagDataList(tag)
        table.getRows().insertByIndex(1, len(dataNodes)-1)  # -1 because there is one already

        for i, node in enumerate(dataNodes):
            try:
                self._writeToTableCell(table, 0, i+1, node.data.ident)
                self._writeToTableCell(table, 1, i+1, node.pathString)
                self._writeToTableCell(table, 2, i+1, node.data.text)
            except:
                print("failed to populate cell:", get_traceback())

        if tag.path == '' or tag.path == '#':
            documentTitle = 'QDA Tags'
        else:
            documentTitle = 'Tag: '+tag.path

        try:
            field = report.getTextFieldMasters().getByName('com.sun.star.text.fieldmaster.User.MAINTAG')
            field.setPropertyValue('Content', documentTitle)
        except:
            print("error: failed to write title:", get_traceback())

        return report

    def _createTagFiltered(self, tag):
        report = self._openDocumentAsNew(self.document)

        if report is None:
            print("error: failed to create new document")
            return

        for field in report.getTextFields():
            if not field.supportsService("com.sun.star.text.TextField.Annotation"):
                continue  # field is not a comment

            tagPath = tag.path
            offset = len(tag.path)

            allTags = sorted([str(x).lower() for x in self.FIND_TAGS_RE.findall(field.Content.strip())])  # e.g. ['#tag1#nested1#nested2', '#tag2']
            matchedTags = [x for x in allTags if x.startswith(tagPath) and (x[offset:] == '' or x[offset:][0] == '#')]

            if not matchedTags:
                # NOTE This removes all regular comments without tags as well.
                # For the time being, this is considered a feature.
                # TODO Do not remove regular comments in comment threads connected
                # to comments that are being kept.
                # Problem: I have found no way to either get a list of text fields
                # at a certain position or to compare anchor/cursor positions.
                # "Replies" to comments are anchored at the same position as the
                # "parent" comment.
                report.getText().removeTextContent(field)

        return report

    def _createTagExport(self, tag):
        report = self._createTagFiltered(tag)

        # collect tagged comments
        # NOTE Currently, all non-tag comments are removed in _createTagFiltered.
        collectedFields = [x for x in report.getTextFields() if \
            x.supportsService("com.sun.star.text.TextField.Annotation") and \
                self.FIND_TAGS_RE.search(x.Content)]

        # prepare color palette
        # We create enough colors that all fields could get
        # a unique color to make sure each unique combination of tags gets its
        # own color.
        #
        # References:
        # - https://seaborn.pydata.org/tutorial/color_palettes.html#using-circular-color-systems
        # - https://www.hsluv.org/
        # - https://stackoverflow.com/a/876872
        # - https://docs.python.org/3/library/colorsys.html
        # - https://github.com/hsluv/hsluv-python

        # TODO first color is too dark and brownish
        count = len(collectedFields)
        hsv_tuples = [(x*1.0/count, 0.5, 0.9) for x in range(count)]     # create equally spread HSV values
        rgb_tuples = map(lambda x: colorsys.hsv_to_rgb(*x), hsv_tuples)  # convert them to RGB
        hpluv_tuples = [hsluv.rgb_to_hpluv(x) for x in rgb_tuples]       # convert RGB to HPLUV (pastel, human-adapted intensity)
        rgb_tuples = [hsluv.hpluv_to_rgb(x) for x in hpluv_tuples]       # convert HPLUV back to RGB
        colors = [int(x[2]*255)+(int(x[1]*255)<<8)+(int(x[2]*255)<<16) for x in rgb_tuples]  # convert RGB tuples to LibreOffice colors

        # replace comments with "baked" color highlights
        # TODO Make sure all highlights are always visible.
        for i, field in enumerate(collectedFields):
            cursor = report.getText().createTextCursorByRange(field.getAnchor())
            cursor.collapseToStart()

            # TODO Insert "[ID]" or "[ID,ID,ID...]"
            report.getText().insertString(cursor, "[??]", True)
            cursor.setPropertyValue('CharWeight', FW_BOLD)
            cursor.setPropertyValue('CharBackColor', colors[i])

            cursor = report.getText().createTextCursorByRange(field.getAnchor())
            cursor.setPropertyValue('CharBackColor', colors[i])

            report.getText().removeTextContent(field)

        # TODO:
        #   - get new document with filtered annotations
        #   - highlight text with different colors, insert tag IDs
        #   - remove all annotations
        #   - -> "<[1] lorem> ipsum dolor <[2] sit <[1,3] amet>, consectetuer> adipiscing elit."

    # --------- helpers ---------------------

    def messageBox(self, MsgText, MsgTitle, MsgType=MESSAGEBOX, MsgButtons=BUTTONS_OK):
        sm = self.LocalContext.ServiceManager
        si = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.LocalContext)
        mBox = si.createMessageBox(self.Toolkit, MsgType, MsgButtons, MsgTitle, MsgText)
        return mBox.execute()

    def _printObjectProperties(self, obj, title=''):
        print("OBJECT PROPERTIES:", title)

        if hasattr(obj, 'getPropertySetInfo'):
            print('\n'.join([f'{x.Name}: {x.Type}' for x in obj.getPropertySetInfo().getProperties()]))
        else:
            print('-> object has no getPropertySetInfo() method')

    def _openDocumentAsNew(self, document):
        '''
        DOES T^emporarily save a document to a temporary location, then open it as a
             new unnamed document and remove the temporary file.
        GETS an opened document
        RETURNS the new document or None
        '''
        try:
            tmpfile = tempfile.NamedTemporaryFile(delete=False, suffix='.odt')
            tmpfile.close()
            tmpfileUrl = 'file://'+tmpfile.name
        except:
            print("error: failed to create temporary file", get_traceback())
            return None

        try:
            document.lockControllers()
            document.storeToURL(tmpfileUrl, ())
            document.unlockControllers()
        except:
            print("error: failed to save document to temporary location", get_traceback())
            os.remove(tmpfile.name)
            return None

        try:
            loadProperties = [PropertyValue("AsTemplate", 0, True, DIRECT_VALUE)]
            newDocument = self.desktop.loadComponentFromURL(tmpfileUrl, "_blank", 0, tuple(loadProperties))
            newDocument.CurrentController.getViewCursor().gotoStart(False)
            os.remove(tmpfile.name)
        except:
            print("error: failed to load and remove temporary file")
            print("-> tempfile:", tempfile.name)
            print("-> traceback:\n", get_traceback())
            return None

        return newDocument

    def _showContextMenu(self, node):
        self._contextMenuItems = {  # inefficient but easier to change
                'rootNode': ['Export all tags', 'Create comprehensive report'],
                'dataNode': ['Move', 'Delete'],
                'tagNode': ['Edit', 'Filter this tag', 'Export this tag', 'Create report', 'Delete'],
            }
        data = self._objectsCache[node.DataValue]

        if isinstance(data, Leaf):
            kind = 'dataNode'
        elif isinstance(data, Tree):
            if not data.path or data.path == '#':
                kind = 'rootNode'
            else:
                kind = 'tagNode'
        else:
            print("warning: requested context menu on unknown node", node)
            return

        self._createContextMenu(kind)
        if not self._contextMenu:
            return

        comp = self.document.getCurrentController().getFrame().getComponentWindow()

        # position of the popupmenu is not considered
        if n := self._contextMenu.execute(comp, Rectangle(), 0):
            print(f"- selected: {n} -> {self._contextMenuItems[kind][n-1]}")

            if kind == 'dataNode':
                if n == 1:
                    print("not yet implemented!")
                    # TODO: move
                    #   - change tag
                    #   - replace tag in all annotations
                elif n == 2:
                    ret = self.messageBox(f'Are you sure you want to delete the tagging of "{node.getDisplayValue()}"?',
                                          'Confirm', WARNINGBOX, BUTTONS_OK_CANCEL)
                    if ret == 1:  # accepted
                        print("not yet implemented!")
            elif kind == 'rootNode':
                if n == 1:
                    self._createTagExport(data)
                elif n == 2:
                    self._createTagReport(data)
            elif kind == 'tagNode':
                if n == 1:
                    print("not yet implemented!")
                    # TODO: edit:
                    #   - open dialog, edit name and description
                    #   - -> description must be saved somewhere, optimally in an external codebook
                elif n == 2:
                    self._createTagFiltered(data)
                elif n == 3:
                    self._createTagExport(data)
                elif n == 4:
                    self._createTagReport(data)
                elif n == 5:
                    ret = self.messageBox(f'Are you sure you want to delete the tag "{node.getDisplayValue()}" '+
                                           'and all information associated with it?', 'Confirm',
                                           WARNINGBOX, BUTTONS_OK_CANCEL)
                    if ret == 1:  # accepted
                        print("not yet implemented!")

    def _createContextMenu(self, kind):
        if kind in self._contextMenuContainer:
            self._contextMenu = self._contextMenuContainer[kind]
            return

        print("building context menu...")
        smgr = self.ctx.getServiceManager()
        popup = smgr.createInstanceWithContext("com.sun.star.awt.PopupMenu", self.ctx)

        if kind not in self._contextMenuItems:
            self.messageBox(f"Bug: unknown context menu '{kind}'", "Error", ERRORBOX)
            return

        for i, item in enumerate(self._contextMenuItems[kind]):
            print(f"- adding menu item #{i+1}: {item}")
            popup.insertItem(i+1, item, 0, i+1)

        self._contextMenuContainer[kind] = popup
        self._contextMenu = popup

    def _scrollToRange(self, textRange):
        '''
        DOES: scrolls the document to the range in a given text
        GETS: a document and a range object. The range must be in the document
        RETURNS: nothing, side effect
        '''
        viewCursor = self.document.CurrentController.getViewCursor()
        viewCursor.gotoRange(textRange, False)

        # It is not necessary to collapse (clear) the selection. Accidental
        # keystrokes only reach the currently focussed item, i.e. the side panel.
        ### viewCursor.collapseToEnd()

    # -----------------------------------------------------------
    #               Execute dialog
    # -----------------------------------------------------------

    def showDialog(self):
        # https://wiki.openoffice.org/wiki/Documentation/DevGuide/GUI/Displaying_Dialogs
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
        selection = self._objectsCache[ev.Source.getSelection().DataValue]

        if not isinstance(selection, Leaf):
            return  # not a data node

        if self.CheckboxJumpto.State == True:
            self._scrollToRange(selection.data.textField.getAnchor())

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

    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/XDocumentEventListener.html
    # https://www.openoffice.org/api/docs/common/ref/com/sun/star/document/DocumentEvent.html
    def notifyEvent(self, ev):
        pass
        # if ev.EventName == 'OnFocus':
            # self.updateTree()
        # print("DOCUMENT EVENT:", ev.EventName)


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
