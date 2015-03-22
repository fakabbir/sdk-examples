import sys
import atexit
import os
import time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


OPENOFFICE_PORT = 2002
_oopaths=(
    ('/usr/lib/libreoffice/program/', '/usr/lib/libreoffice/program/'),
    ('/usr/lib64/libreoffice/program/', '/usr/lib64/libreoffice/program/'),
)

for p in _oopaths:
    if os.path.exists(p[0]):
        OPENOFFICE_PATH    = p[0]
        OPENOFFICE_BIN     = os.path.join(OPENOFFICE_PATH, 'soffice')
        OPENOFFICE_LIBPATH = p[1]

        # Add to path so we can find uno.
        if sys.path.count(OPENOFFICE_LIBPATH) == 0:
            sys.path.insert(0, OPENOFFICE_LIBPATH)
        break

class OORunner:
    """
    Start, stop, and connect to OpenOffice.
    """
    def __init__(self, port=OPENOFFICE_PORT):
        """ Create OORunner that connects on the specified port. """
        self.port = port

    def connect(self, no_startup=False):
        """
        Connect to OpenOffice.
        If a connection cannot be established try to start OpenOffice.
        """
        localContext = uno.getComponentContext()
        servicemanager = localContext.ServiceManager
        resolver = servicemanager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        context = None
        did_start = False

        n = 0
        while n < 6:
            try:
                context = resolver.resolve("uno:socket,host=localhost,port=%d;urp;StarOffice.ComponentContext" % self.port)
                break
            except NoConnectException:
                pass

            # If first connect failed then try starting OpenOffice.
            if n == 0:
                # Exit loop if startup not desired.
                if no_startup:
                     break
                self.startup()
                did_start = True

            # Pause and try again to connect
            time.sleep(1)
            n += 1

        if not context:
            raise Exception("Failed to connect to OpenOffice on port %d" % self.port)

        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        # generate a graphic provider as well to embed images from filesystem
        graphicprovider = context.ServiceManager.createInstance('com.sun.star.graphic.GraphicProvider')

        if not (desktop and graphicprovider):
            raise Exception("Failed to create OpenOffice desktop or graphicprovider on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop
            _started_graphicproviders[self.port] = graphicprovider

        return desktop, graphicprovider

    def startup(self):
        """
        Start a headless instance of OpenOffice.
        """
        args = [OPENOFFICE_BIN,
                '--headless',
                '--accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager;tcpNoDelay=1;' % self.port,
                '--nologo',
                '--invisible',
                '--nofirststartwizard',
                '--nocrashreport',
                '--norestore'
                ]
        env  = {'PATH'       : '/bin:/usr/bin:%s' % OPENOFFICE_PATH,
                'PYTHONPATH' : OPENOFFICE_LIBPATH,
                }

        try:
            pid = os.spawnve(os.P_NOWAIT, args[0], args, env)
        except Exception as e:
            raiseException("Failed to start OpenOffice on port %d: %s" % (self.port, e.message))

        if pid <= 0:
            raise Exception("Failed to start OpenOffice on port %d" % self.port)

    def shutdown(self):
        """
        Shutdown OpenOffice.
        """
        try:
            if _started_desktops.get(self.port):
                _started_desktops[self.port].terminate()
                del _started_desktops[self.port]
            if _started_graphicproviders.get(self.port):
                _started_graphicproviders[self.port].terminate()
                del _started_graphicproviders[self.port]
        except Exception as e:
            pass

# Keep track of started desktops and shut them down on exit.
_started_desktops = {}
_started_graphicproviders = {}

def _shutdown_desktops():
    """ Shutdown all OpenOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
        try:
            if desktop:
                desktop.terminate()
        except Exception as e:
            pass

def _shutdown_graphicproviders():
    """ Shutdown all OpenOffice graphicproviders that were started by the program. """
    for port, graphicprovider in _started_graphicproviders.items():
        try:
            if graphicprovider:
                graphicprovider.terminate()
        except Exception as e:
            pass


atexit.register(_shutdown_desktops)
atexit.register(_shutdown_graphicproviders)


def oo_shutdown_if_running(port=OPENOFFICE_PORT):
    """ Shutdown OpenOffice if it's running on the specified port. """
    oorunner = OORunner(port)
    try:
        desktop, graphicprovider = oorunner.connect(no_startup=True)
        desktop.terminate()
        graphicprovider.terminate()
    except Exception as e:
        pass


def oo_properties(**args):
    """
    Convert args to OpenOffice property values.
    """
    props = []
    for key in args:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)

    return tuple(props)