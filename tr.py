#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# This file incorporates work covered by the following license notice:
#
#   Licensed to the Apache Software Foundation (ASF) under one or more
#   contributor license agreements. See the NOTICE file distributed
#   with this work for additional information regarding copyright
#   ownership. The ASF licenses this file to you under the Apache
#   License, Version 2.0 (the "License"); you may not use this file
#   except in compliance with the License. You may obtain a copy of
#   the License at http://www.apache.org/licenses/LICENSE-2.0 .
#

import re
import uno
import json
import sys
from googletrans import Translator

# helper function
def getNewString( theString ) :
    translator = Translator()
    if( not theString or len(theString) ==0) :
        return ""
    # should we tokenize on "."?
    else : 
        #newString = re.sub('\n',' ', theString.rstrip())
        newString=translator.translate( theString  , dest='uk')
    return newString.text

def cP( ): 
    """Change the case of a selection, or current word from upper case, to first char upper case, to all lower case to upper case..."""
    import string
    # The context variable is of type XScriptContext and is available to
    # all BeanShell scripts executed by the Script Framework
    xModel = XSCRIPTCONTEXT.getDocument()
    #the writer controller impl supports the css.view.XSelectionSupplier interface
    xSelectionSupplier = xModel.getCurrentController()
    #see section 7.5.1 of developers' guide
    xIndexAccess = xSelectionSupplier.getSelection()
    count = xIndexAccess.getCount();
    if(count>=1):  #ie we have a selection
        i=0
    while i < count :
            xTextRange = xIndexAccess.getByIndex(i);
            #print "string: " + xTextRange.getString();
            theString = xTextRange.getString();
            if len(theString)==0 :
                # sadly we can have a selection where nothing is selected
                # in this case we get the XWordCursor and make a selection!
                xText = xTextRange.getText();
                xWordCursor = xText.createTextCursorByRange(xTextRange);
                if not xWordCursor.isStartOfWord():
                    xWordCursor.gotoStartOfWord(False);
                xWordCursor.gotoNextWord(True);
                theString = xWordCursor.getString();
                newString = getNewString(theString);
                if newString :
                    xWordCursor.setString(newString);
                    xSelectionSupplier.select(xWordCursor);
            else :

                newString = getNewString( theString );
                if newString:
                    xTextRange.setString(newString);
                    xSelectionSupplier.select(xTextRange);
            i+= 1


# lists the scripts, that shall be visible inside OOo. Can be omitted, if
# all functions shall be visible, however here getNewString shall be suppressed
g_exportedScripts = cP,

# vim: set shiftwidth=4 softtabstop=4 expandtab:
