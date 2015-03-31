import sys
import atexit
import os
import time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException


LIBREOFFICE_PORT = 2002
_libreoffice_paths=(
    ('/usr/lib/libreoffice/program/', '/usr/lib/libreoffice/program/'),
    ('/usr/lib64/libreoffice/program/', '/usr/lib64/libreoffice/program/'),
)

for p in _libreoffice_paths:
    if os.path.exists(p[0]):
        LIBREOFFICE_PATH    = p[0]
        LIBREOFFICE_BIN     = os.path.join(LIBREOFFICE_PATH, 'soffice')
        LIBREOFFICE_LIBPATH = p[1]

        # Add to path so we can find uno.
        if sys.path.count(LIBREOFFICE_LIBPATH) == 0:
            sys.path.insert(0, LIBREOFFICE_LIBPATH)
        break

class LibreOfficeRunner:
    """
    Start, stop, and connect to LibreOffice.
    """
    def __init__(self, port=LIBREOFFICE_PORT):
        """ Create LibreOfficeRunner that connects on the specified port. """
        self.port = port

    def connect(self, no_startup=False):
        """
        Connect to LibreOffice.
        If a connection cannot be established try to start LibreOffice.
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

            # If first connect failed then try starting LibreOffice.
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
            raise Exception("Failed to connect to LibreOffice on port %d" % self.port)

        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        # generate a graphic provider as well to embed images from filesystem
        graphicprovider = context.ServiceManager.createInstance('com.sun.star.graphic.GraphicProvider')

        if not (desktop and graphicprovider):
            raise Exception("Failed to create LibreOffice desktop or graphicprovider on port %d" % self.port)

        if did_start:
            _started_desktops[self.port] = desktop
            _started_graphicproviders[self.port] = graphicprovider

        return desktop, graphicprovider

    def startup(self):
        """
        Start a headless instance of LibreOffice.
        """
        args = [LIBREOFFICE_BIN,
                '--headless',
                '--accept=socket,host=localhost,port=%d;urp;StarOffice.ServiceManager;tcpNoDelay=1;' % self.port,
                '--nologo',
                '--invisible',
                '--nofirststartwizard',
                '--nocrashreport',
                '--norestore'
                ]
        env  = {'PATH'       : '/bin:/usr/bin:%s' % LIBREOFFICE_PATH,
                'PYTHONPATH' : LIBREOFFICE_LIBPATH,
                }

        try:
            pid = os.spawnve(os.P_NOWAIT, args[0], args, env)
        except Exception as e:
            raiseException("Failed to start LibreOffice on port %d: %s" % (self.port, e.message))

        if pid <= 0:
            raise Exception("Failed to start LibreOffice on port %d" % self.port)

    def shutdown(self):
        """
        Shutdown LibreOffice.
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
    """ Shutdown all LibreOffice desktops that were started by the program. """
    for port, desktop in _started_desktops.items():
        try:
            if desktop:
                desktop.terminate()
        except Exception as e:
            pass

def _shutdown_graphicproviders():
    """ Shutdown all LibreOffice graphicproviders that were started by the program. """
    for port, graphicprovider in _started_graphicproviders.items():
        try:
            if graphicprovider:
                graphicprovider.terminate()
        except Exception as e:
            pass


atexit.register(_shutdown_desktops)
atexit.register(_shutdown_graphicproviders)


def oo_shutdown_if_running(port=LIBREOFFICE_PORT):
    """ Shutdown LibreOffice if it's running on the specified port. """
    oorunner = LibreOfficeRunner(port)
    try:
        desktop, graphicprovider = oorunner.connect(no_startup=True)
        desktop.terminate()
        graphicprovider.terminate()
    except Exception as e:
        pass


def oo_properties(**args):
    """
    Convert args to LibreOffice property values.
    """
    props = []
    for key in args:
        prop = PropertyValue()
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)

    return tuple(props)
