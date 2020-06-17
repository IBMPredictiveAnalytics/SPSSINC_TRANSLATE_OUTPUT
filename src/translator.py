#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2020
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

# pivot table translation framework
# translations are expected to be structed as ini files suitable for the SafeConfigParser class in ConfigParser
# The keys are the label text items in the pivot tables.  The values are the translations of these items.

# structure consists of a globals file named GLOBALTRANS.ini, an optional locals file named LOCALTRANS.ini,
# and optional locals files with names of the form OMS-subtype.ini

# globals items are used as a fallback if an entry is not found in the locals file or OMS-specific file
# for the current table subtype
# globals must contain one section named GLOBALS.  Any other sections are ignored

# locals file contains 0 or more sections with names matching the OMS table subtype.
# The section names, except for GLOBALS, should be written in lower case with no blanks.
# Each section, in a locals file may contain a special entry
# TSCOPE={ROWS | COLUMNS | ALL*}
# indicating whether to translate only row labels, only column labels, or all labels
#
# Both LOCALTRANS and the OMS-named files have table-specific contents.  Usually a set of separate files
# or a single LOCALTRANS file with multiple sections would be used.  Which to use is the translator's choice.


# Entries in the ini files are of the form
# name=value
# name is the text of the label in a table to be replaced by a translation.

# The ini files must be in UTF_8 with a (utf_8) BOM.
# By default, the translation files are expected to be in the extensions subdirectory of the PASW Statistics
# installation.  However, if the environment variable SPSS_TRANSLATOR is defined, the files are expected
# to be there.

#The translation for outline description and heading items must be in the GLOBALS file.

# Text with parameters can be translated using "regexp" sections related to the standard sections.
# See the document accompanying this module for details.

# This mechanism is meant to be used as a base or type-specific autoscript, or it can be used
# to translate "offline", processing all the pivot tables and other translatable items in a Viewer.

# Note: footnote translation does not work in version 19.0.0 if there is more than one footnote.

version = "1.1.0"
author = "SPSS, JKP"

# history
# 11-May-2009 initial version
# 21-Sep-2009 partial support for translating layer labels
# 26-Aug-2010 support for table footnote translation


import configparser, codecs, re, inspect, os.path
from os import getenv
import SpssClient

class Translator(object):
    """translate pivot tables, titles, and descriptions"""
    def __init__(self, folder=None):
        """If not None, folder specifies the folder where the translation definition files will be found.
        Otherwise these files are found from the SPSS_TRANSLATOR environment variable or
        the extensions subdirectory of the product installation."""

        self.tc = None
        if folder is None:
            self.tpath = getenv("SPSS_TRANSLATOR")
            if self.tpath is None:
                self.tpath= SpssClient.GetSPSSPath().replace("\\", "/") + "/extensions"
            else:
                self.tpath.rstrip("\\/")
        else:
            self.tpath = folder.rstrip("\\/")
        self.alreadyread = set()  # used to track which translation files have been loaded.
        # self.regexps will hold a list of compiled regular expressions with the section name as the key
        # regexp sections follow the name of the ordinary section with suffix "-regexp"
        # self.noregexp is a set indicating whether an attempt has been made to get a regexp section
        # for this section in order to avoid repeated attempts to load one
        self.regexps = {}
        self.noregexp = set()
        
    
    def replaceText(self, subtype, getter, setter, indexes=None):
        """lookup translation and change text if there is a translation.
        
        subtype is the table subtype or "" when none
        getter is the bound method for retrieving the text to translate
        setter is the bound method for setting the translated text
        If given, indexes is a duple of the row and column numbers or a single number for layers.  Applies only to pivot tables.
        
        If local and global lookups fail, an attempt is made to substitute from a regexp section if one exists."""
        
        trans = None
        if indexes is None:
            text = getter()
        else:
            numsub = len(indexes)
            if numsub == 2:
                text = getter(indexes[0], indexes[1])
            else:
                text = getter(indexes[0])
        try:
            trans = self.tc.get(subtype, text, raw=True)
        except:
            try:
                trans = self.tc.get("GLOBALS", text, raw=True)
            except:
                trans = self.doregexp(subtype, text)

        if not trans is None:
            if indexes is None:
                try:   # SetDimensionName api does not currently accept extended characters
                    setter(trans)
                except:
                    pass
            else:
                try:
                    if numsub == 2:
                        setter(indexes[0], indexes[1], trans)
                    else:
                        setter(indexes, 0, trans)
                except:
                    pass    #v17 will sometimes fail on the last item in a dimension
                
    def doregexp(self, subtype, text):
        """try to translate via applicable regular expression compiling as we go.
        
        subtype is the section
        text is the text to translate
        
        The first matching regular expression wins.
        No scoping is applied here."""

        for sec in (subtype, "GLOBALS"):
            if sec in self.noregexp:   # known that there are no re's for this section
                continue
            if sec not in self.regexps:
                # try to load and compile a relevant section of re's
                relist = []
                try:
                    for pat, repl in self.tc.items(sec + "-regexp", raw=True):
                        relist.append((re.compile(pat, re.LOCALE), repl))
                    self.regexps[sec] = relist
                except:
                    # no regexp section
                    self.noregexp.add(sec)
            if sec in self.regexps:
                for regexp in self.regexps[sec]:
                    retrans, subcount = re.subn(regexp[0], regexp[1], text)
                    if subcount > 0:
                        return retrans
        return None
    
    def translateItem(self, item):
        """translate a pivot table or other item if possible.
        
        item is a pivot table, description, title, or heading item to translate"""
        
        itemtype = item.GetType()            
        if not itemtype in [SpssClient.OutputItemType.PIVOT, SpssClient.OutputItemType.TITLE, 
            SpssClient.OutputItemType.HEAD]:
            return
    
        if self.tc is None:
            # load the translations
            self.tc = configparser.SafeConfigParser()
            self.tc.optionxform = lambda x: x   # translation keys are case sensitive
        if itemtype == SpssClient.OutputItemType.PIVOT:
            subtype = item.GetSubType()
            subtype = "".join(subtype.lower().split())
        else:
            subtype = ""
        # Ensure that relevant translation definitions are loaded
        for i, fname in enumerate(["/GLOBALTRANS.ini", "/LOCALTRANS.ini", "/%s.ini" % subtype]):
            if i == 2 and subtype == "":
                break
            try:
                if not fname in self.alreadyread:
                    fp = codecs.open(self.tpath + fname, "rU", encoding="utf_8_sig")
                    self.tc.readfp(fp)
                    self.alreadyread.add(fname)
            except:
                pass
        if not self.alreadyread:        # do we have something to do?
            return
        if itemtype == SpssClient.OutputItemType.HEAD:
            self.replaceText(subtype, item.GetDescription, item.SetDescription)
        elif itemtype == SpssClient.OutputItemType.TITLE:
            spitem = item.GetSpecificType()
            self.replaceText(subtype, spitem.GetTextContents, spitem.SetTextContents)
        elif subtype != "":   # is this a pivot table?
            spitem = item.GetSpecificType()
            try:
                tscope = self.tc.get(subtype, "TSCOPE")
                rows = tscope.lower() in ["rows", "all"]
                cols = tscope.lower() in ["columns", "all"]
            except:
                rows, cols = True, True
            
            # TODO: deal with layers.
            spitem.SetUpdateScreen(False)
            while rows or cols:
                if rows:
                    rows = False
                    labels = spitem.RowLabelArray()
                elif cols:
                    cols = False
                    labels = spitem.ColumnLabelArray()
                    
                # try to translate label items
                for r in range(labels.GetNumRows()):
                    for c in range(labels.GetNumColumns()):
                        self.replaceText(subtype, labels.GetValueAt, labels.SetValueAt, (r,c))
                self.replaceText(subtype, spitem.GetTitleText, spitem.SetTitleText)
                self.replaceText(subtype, spitem.GetCaptionText, spitem.SetCaptionText)
                
                # try to translate footnotes
                fnarray = spitem.FootnotesArray()
                def fngetter(i):
                    """wrapper for footnote text retriever to eliminate leading blanks
                
                    i is the index of the footnote to retrieve"""
                    return fnarray.GetValueAt(i).lstrip().rstrip("\n")

                def fnsetter(indexes, i, text):
                    """wrapper for footnote text setter
                    
                    indexes is the array of indexess
                    i is the index into the array
                    text is the text to set"""
                    
                    fnarray.SetValueAt(indexes[i], text)
                
                for i in range(fnarray.GetCount()):
                    self.replaceText(subtype, fngetter, fnsetter,[i])
                
                
                # work on layers - the apis are a bit different
                layerlabel = spitem.LayerLabelArray()
                numlayerdims = layerlabel.GetNumDimensions()
                if numlayerdims > 0:
                    pm = spitem.PivotManager()
                for layer in range(numlayerdims):
                    layerdim = pm.GetLayerDimension(layer)
                    self.replaceText(subtype, layerdim.GetDimensionName, layerdim.SetDimensionName, None)
                    #numcats = layerdim.GetNumCategories()
                    #for c in range(numcats):
                        #self.replaceText(subtype, layerdim.GetCategoryValueAt, layerdim.SetCategoryValueAt, (c,))
                    
                
            # description is a property of the non-specific item
            self.replaceText(subtype, item.GetDescription, item.SetDescription)
            spitem.SetUpdateScreen(True)
    
def dotrans(folder=None, importing=False, process="preceding", selectedonly=False, subtype=["*"]):

    SpssClient.StartClient()
    context = SpssClient.GetScriptContext()
    
    # if this is not an autoscript, it is triggered by the extension command, but we
    # don't want the import to do anything.  The script is run directly only as an autoscript.
    if importing and not context:
        return

    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
        #import thread
        #wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
        ## for V19 use
        ##    ###SpssClient._heartBeat(False)
    #except:
        #pass
    t = Translator(folder)
    if context is not None:     # are we an autoscript?
        item = context.GetOutputItem()  # which item?
        t.translateItem(item)
    else:   # Not an autoscript
        items = SpssClient.GetDesignatedOutputDoc().GetOutputItems()
        itemcount = items.Size()
        if "*" in subtype:
            subtype = ["*"]
        else:
            # remove white space - subtype must already be a sequence
            subtype = ["".join(st.lower().split()) for st in subtype]
            # remove matching outer quotes of any type
            subtype = [re.sub(r"""^('|")(.*)\1$""", r"""\2""", st) for st in subtype]

        if items.GetItemAt(itemcount-1).GetType() == SpssClient.OutputItemType.LOG:
            itemcount -= 1
        for itemnumber in range(itemcount-1, -1, -1):
            item = items.GetItemAt(itemnumber)
            if selectedonly and not item.IsSelected():
                continue
            if not selectedonly and process == "preceding" and item.GetTreeLevel() <= 1:
                break
            if item.GetType() == SpssClient.OutputItemType.PIVOT:
                if subtype[0] != "*" and not "".join(item.GetSubType().lower().split()) in subtype:
                    continue
            t.translateItem(item)
                    
    SpssClient.StopClient()
    
try:
    cc2 = inspect.stack()[1][1]
except:
    cc2 = ""
if os.path.splitext(os.path.basename(cc2))[0] != 'SPSSINC_TRANSLATE_OUTPUT':
    dotrans(importing=True)
