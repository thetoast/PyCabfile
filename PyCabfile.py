"""
A lightweight ctypes library for accessing Cabinet files.

This is a python module which wraps the Windows SetupAPI for accessing and extracting
files from within a Microsoft Cabinet File

See the Cabinet class for details on using this module.

Note: At this time, the module only supports extracting files from a cabinet.  It does
not support modifying the cabinet in any way.

Changelog:
    4 Jan 2011 - 1.0:
        * Initial version with read-only support of cabfiles.
        * Only supports extracting of files.

"""

__version__ = "1.0"
__author__ = "Ryan Mechelke <rfmechelke@gmail.com>"
__license__ = "Public Domain"
__date__ = "4 Jan 2011"

__all__ = "Cabinet, CabinetFile, CabinetError"

from ctypes import *
from glob import glob
from os import getcwd, mkdir, stat, sep

# constants for relevant parts of the Windows Setup API
SPFILENOTIFY_CABINETINFO = 0x00000010
SPFILENOTIFY_FILEINCABINET = 0x00000011
SPFILENOTIFY_NEEDNEWCABINET = 0x00000012
SPFILENOTIFY_FILEEXTRACTED = 0x00000013
FILEOP_SKIP = 2
FILEOP_DOIT = 1
FILEOP_ABORT = 0
NO_ERROR = 0
MAX_PATH = 260


# ctypes callback needed for SetupIterateCabinet
PSP_FILE_CALLBACK = WINFUNCTYPE(c_uint, c_void_p, c_uint, POINTER(c_uint), POINTER(c_uint))


# ctypes structure for getting data about compressed files through PSP_FILE_CALLBACK
class FILE_IN_CABINET_INFO(Structure):
    _fields_ = [
        ("NameInCabinet", c_wchar_p),
        ("FileSize", c_ulong),
        ("Win32Error", c_ulong),
        ("DosDate", c_ushort),
        ("DosTime", c_ushort),
        ("DosAttribs", c_ushort),
        ("FullTargetName", c_wchar * MAX_PATH)
    ]


class CabinetError(Exception):
    """
    This is an Exception class which is used to represent problems when
    working with cabinets and their contained files.
    """
    
    def __init__(self, message=None, err=None):
        if err:
            message = "Error 0x%X: %s" % (err, FormatError(err))

        super(Exception, self).__init__(message)


class CabinetFile(object):
    """
    This class represents a file which is contained within a cabinet.
    """

    def __init__(self, cabinet, name):
        """
        Creates a new CabinetFile.  Not to be used directly.
        """

        self.cabinet = cabinet
        self.name = name

    def __str__(self):
        return str(self.name)

    def __unicode__(self):
        return unicode(self.name)

    def __repr__(self):
        return unicode(self)

    def extract(self, dest=None):
        """
        Extracts the file from its Cabinet.

        Essentially, this just calls self.cabinet.extract(self.name, dest).
        """

        return self.cabinet.extract(self.name, dest)


class Cabinet(object):
    """
    This class does all the work of processing microsoft cabinet files.

    Usage:
    Create an instance of the class thusly:

        mycab = Cabinet("path_to_cab")
    
    Now, contained files can be browsed as follows:
    
        for cabfile in mycab.files:
            print cabfile

    Individual files can be extracted by doing:
        
        myfile = mycab.files[0]
        myfile.extract()

        - or -

        mycab.extract('SomeFile.txt')

    All files can be extracted as well:

        mycab.extract_all()

    Note that each of these extract methods can take optional destination paths.
    
    """

    def __init__(self, cabinet_path):
        """
        Creates a new Cabinet instance wrapping the provided cab file.
        """
        self.name = unicode(cabinet_path)
        self.files = []
        self._extract_all = False
        self._file_to_extract = None
        self._dest = None

        self._list_files = True
        self._do_callback()
        self._list_files = False

    def _do_callback(self):
        file_callback = PSP_FILE_CALLBACK(self._py_file_callback)

        retval = windll.setupapi.SetupIterateCabinetW(self.name, 0, file_callback, None)

        if not retval:
            raise CabinetError(err=GetLastError())

    def _py_file_callback(self, context, notification, param1, param2):

        if notification == SPFILENOTIFY_FILEINCABINET:
            file_info_p = cast(param1, POINTER(FILE_IN_CABINET_INFO))
            file_name = file_info_p.contents.NameInCabinet

            if self._list_files:
                _file = CabinetFile(self, file_name)
                _file._file_in_cab = file_info_p.contents
                self.files.append(_file)
            elif self._extract_all:
                last_part = unicode(file_name[file_name.rfind(u"\\") + 1:])
                extract_path = self._dest + u"\\" + last_part + u"\u0000"

                file_info_p.contents.FullTargetName = extract_path

                return FILEOP_DOIT
            elif file_name == self._file_to_extract:
                file_info_p.contents.FullTargetName = self._dest + u"\u0000"

                return FILEOP_DOIT
                
            return FILEOP_SKIP
        elif notification == SPFILENOTIFY_CABINETINFO: pass
        elif notification == SPFILENOTIFY_NEEDNEWCABINET: pass
        elif notification == SPFILENOTIFY_FILEEXTRACTED: pass
        else: raise CabinetError(u"Unknown notification type: %X" % (notification))

        return NO_ERROR

    def _getdir(self):
        cabfile = self.name
        dirname = cabfile[cabfile.rfind(u"\\") + 1:cabfile.find(u".")]
        try:
            stat(dirname)
        except WindowsError:
            mkdir(dirname)

        return dirname

    def extract(self, file_name, dest=None):
        """
        Extracts a file from the cabinet to the location specified by dest.

        If dest is specified, and is a directory, the file will be extracted
        to the given directory and the file name will be retained.

        If dest is specified and does not point to an existing directory, the
        Cabinet object will extract the file to the exact dest specified.

        If dest is not specified, this will create a directory in the
        current working directory which is named after the cab file,
        and then it will extract the given file there, retaining the existing
        file name.

        This method returns the location of the newly extracted file.
        """

        if not dest:
            uni_file_name = unicode(file_name)
            last_part = uni_file_name[uni_file_name.rfind(u"\\") + 1:]
            dest = unicode(getcwd()) + u"\\" + self._getdir() +  u"\\" + last_part

        self._file_to_extract = unicode(file_name)
        self._dest = unicode(dest)

        self._do_callback()

        self._file_to_extract = None
        self._dest = None

        return dest

    def extract_all(self, dest=None):
        """
        Extracts all files from the cabinet into the directory identified by dest.

        If dest is specified, but does not exist, it will be created.

        If dest is not specified, this will create a directory in the
        current working directory which is named after the cab file,
        and then it will extract all files there.

        This method returns the location of the newly extracted files.
        """

        if not dest:
            dest = unicode(getcwd()) + u"\\" + self._getdir()
        else:
            try:
                stat(dest)
            except WindowsError:
                mkdir(dest)

        self._extract_all = True
        self._dest = unicode(dest)

        self._do_callback()

        self._extract_all = False
        self._dest = None

        return dest

    def __str__(self):
        return str(self.name)

    def __unicode__(self):
        return unicode(self.name)

    def __repr__(self):
        return unicode(self)
