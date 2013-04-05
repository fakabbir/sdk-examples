# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import datetime

import uno, unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.document import XEventListener


class TuesdayPrinter(unohelper.Base, XJobExecutor, XEventListener):
    """A service only printing on Tuesdays

    Its important that this service implements this interface:
        - http://api.libreoffice.org/docs/common/ref/com/sun/star/task/XJobExecutor.html
    has a trigger function. Because of that, it can be used in menus and toolbars, by using:

        service:org.libreoffice.TuesdayPrinter

    as a URL.
    """
    def trigger(self, args):
        frame = self.desktop.ActiveFrame
        if datetime.date.today().weekday() == 1:
            # its Tuesday, we really call print now using the '.uno:Print' url
            # that was removed from the menus and toolbars
            # see http://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Using_the_Dispatch_Framework for all the details
            self.dispatchhelper.executeDispatch(
                frame,
                ".uno:Print",
                "_self",
                0, ())
        else:
            window = frame.ContainerWindow
            window.Toolkit.createMessageBox(
                window,
                uno.createUnoStruct("com.sun.star.awt.Rectangle"),
                "warningbox",
                uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_OK"),
                "Im sorry, Dave.",
                "Im afraid I cant do that. It is not Tuesday.").execute()
    # boilerplate code below this point
    def __init__(self, context):
        self.context = context
        # see http://api.libreoffice.org/docs/common/ref/com/sun/star/frame/Desktop.html
        self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        # see http://api.libreoffice.org/docs/common/ref/com/sun/star/frame/DispatchHelper.html
        self.dispatchhelper = self.createUnoService("com.sun.star.frame.DispatchHelper")
    def createUnoService(self, name):
        """little helper function to create services in our context"""
        # see http://api.libreoffice.org/docs/common/ref/com/sun/star/lang/ServiceManager.html
        # see http://api.libreoffice.org/docs/common/ref/com/sun/star/lang/XMultiComponentFactory.html#createInstanceWithContext
        return self.context.ServiceManager.createInstanceWithContext(name, self.context)
    def disposing(self, args):
        pass
    def notifyEvent(self, args):
        pass

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    TuesdayPrinter,
    "org.libreoffice.TuesdayPrinter",
    ("com.sun.star.task.JobExecutor",))
