# Construct a dataset listing the variables and selected properties for a collection of data files
from __future__ import with_statement
#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2014
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

# 05-23-2008 Original version - JKP
# 04-29-2009 Add file handle support

__version__ = "1.0.1"
__author__ = "JKP, SPSS"

from extension import Template, Syntax, processcmd
import translator

#try:
    #import wingdbstub ### debug
#except:
    #pass

def Run(args):
    """Execute the SPSSINC TRANSLATE OUTPUT command"""
    
    
    ###print args   #debug
    args = args[args.keys()[0]]
    
    helptext=r"""SPSSINC TRANSLATE OUTPUT [FOLDER=folder-specification]
    [SUBTYPE=list of subtypes]  
    [PROCESS={PRECEDING* | ALL}
    [SELECTEDONLY={NO* | YES}
    [/HELP]
    
    Translate the contents of the designated Viewer for the object types supported.
    
    FOLDER, if specified, defines the folder where the translation definition files are located.
    Otherwise the files are expected to be found based on the SPSS_TRANSLATOR 
    environment variable or in the extensions subdirectory of the installation folder.
    
    SELECTEDONLY = YES causes only the selected items in the Viewer to be translated.
    
    SUBTYPE can specify a list of OMS table subtypes.  If given, only tables of those types
    will be translated.  This is ignored if SELECTEDONLY is YES.
    
    PROCESS specifies whether to process only the immediately preceding procedure output
    or the entire Viewer contents.  This is ignored if SELECTEDONLY is YES.
    
    /HELP displays this text and does nothing else.
    
    Examples:
    SPSSINC TRANSLATE OUTPUT FOLDER="C:/translationfiles" SUBTYPE="Custom Table".
    
    SPSSINC TRANSLATE OUTPUT SELECTEDONLY=YES.
"""
    
    oobj = Syntax([
    Template("FOLDER", subc="",  ktype="literal", var="folder"),
    Template("SUBTYPE", subc="",  ktype="str", var="subtype", islist=True),
    Template("PROCESS", subc="", ktype="str", vallist=["preceding", "all"], var="process"),
    Template("SELECTEDONLY", subc="", ktype="bool", var="selectedonly"),
    Template("HELP", subc="", ktype="bool")])
    
    # A HELP subcommand overrides all else
    if args.has_key("HELP"):
        #print helptext
        helper()
    else:
        processcmd(oobj, args, translator.dotrans)

def helper():
    """open html help in default browser window
    
    The location is computed from the current module name"""
    
    import webbrowser, os.path
    
    path = os.path.splitext(__file__)[0]
    helpspec = "file://" + path + os.path.sep + \
         "markdown.html"
    
    # webbrowser.open seems not to work well
    browser = webbrowser.get()
    if not browser.open_new(helpspec):
        print("Help file not found:" + helpspec)
try:    #override
    from extension import helper
except:
    pass