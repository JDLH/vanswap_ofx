#!/usr/bin/env python2.7
# encoding: utf-8
'''
vanswap_ofx -- swap NAME and MEMO fields in OFX files 

vanswap_ofx is a utility to repair OFX files by swapping the NAME 
and MEMO fields. It works around a problem with OFX files created 
by my credit union after a system upgrade in 2016. They generated OFX
files with the NAME string in the MEMO field, and vice versa. This 
utility repairs the OFX file by reading it in, switching the values
of NAME and MEMO in each transaction, and writing the file out to 
a different filename.

It defines a class OFXRepairer, which does the work of parsing the
OFX file, performing the repair, and writing out the file with a 
modified filename.

@author:     Jim DeLaHunt

@copyright:  Main program in the public domain. Some modules copyright their authors.

@license:    Public domain, with components MIT licenced.

@contact:    info@jdlh.com
@deffield    updated: Updated
'''

import sys
import os.path
import io
import codecs
import re
# Danger, ofxparse.ofxparse and OfxFile are not official exports of ofxparse.
from ofxparse.ofxparse import OfxFile

from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter


__all__ = []
__version__ = 0.1
__date__ = '2016-12-08'
__updated__ = '2016-12-08'

DEBUG = 0
TESTRUN = 1
PROFILE = 0

class CLIError(Exception):
    '''Generic exception to raise and log different fatal errors.'''
    def __init__(self, msg):
        super(CLIError).__init__(type(self))
        self.msg = "E: %s" % msg
    def __str__(self):
        return self.msg
    def __unicode__(self):
        return self.msg
    
class PathFile(object):
    '''PathFile: instantiate with either a fileobj or a path, get a fileobj
    
    PathFile handles opening a path to get a fileobj. 

    (We will use a temp file pathname for examples.)
    >>> import tempfile
    >>> f = tempfile.NamedTemporaryFile(delete=False)
    >>> f.write(''); f.close()
    
    Pass in a string or bytes array for the filename, and PathFile opens it.
    The attribute '_fh' contains the fileobject for the opened file.
    >>> pf = PathFile(f.name, 'r'); hasattr(pf, 'read')
    True
    
    The attribute 'name' contains the path name supplied to PathFile.
    >>> pf.name == f.name
    True
    
    PathFile opens the path on instantiation, and closes on destruction.
    >>> pf.closed
    False
    >>> t_fh = pf._fh; del(pf); t_fh.closed; del(t_fh)
    True
    
    PathFile will raise the same exceptions as open(). For example, if 
    PathFile tries to open a file for reading, and there is no file at the 
    path, PathFile raises an IOError.
    >>> p = f.name; os.remove(p)  # p is path where no file exists
    >>> pf = PathFile(p); pf is None  # Lazy opening: pf is not yet open.
    False
    >>> pf.open()     # doctest: +IGNORE_EXCEPTION_DETAIL
    Traceback (most recent call last):
      ...
    IOError: [Errno 2] No such file or directory: ...
    
    PathFile also accepts an open file-like object. It will use that object
    instead of treating the parameter as a pathname to open.
    >>> import io
    >>> s = io.BytesIO(b'DUMMY Test file contents')
    >>> pf = PathFile(s); pf.read()
    'DUMMY Test file contents'
    
    Supplying file-like objects is helpful in writing test cases. 
    Test cases can pass a BytesIO object with test data to the PathFile object.
    Production clients, on the other hand, can pass a path.
    '''

    def __init__(self, p, mode='r'):
        self._fh = None
        self.name = None
        self._mode = mode
        if self._is_file(p):
            self._fh = p
            try:
                self.name = p.name
            except AttributeError:
                self.name = repr(p)
            return
        # self._fh = open(p, mode)
        self.name = p
        return
    
    def __del__(self):
        '''PathFile.__del__: close fileobj, if open.'''
        if self._fh:
            self._fh.close()

    def _is_file(self, p):
        '''returns True if p is a file-like object, False otherwise
        
        p: an object to examine
        '''
        return hasattr(p, 'read') and hasattr(p, 'close')
    
    # emulate file class's methods by passing all attribute lookups to 
    # _fh, the underlying file object.
    # This emulation based on class tempfile._TemporaryFileWrapper.
    def __getattr__(self, name):
        # Attribute lookups are delegated to the underlying file
        # and cached for non-numeric results
        # (i.e. methods are cached, closed and friends are not)
        _fh = self.__dict__['_fh'] # Can't say self._fh, that would be recursive 
        if _fh is None:
            _fh = self._fh = io.open(self.__dict__['name'], self.__dict__['_mode'])
        a = getattr(_fh, name)
        if not issubclass(type(a), type(0)):
            setattr(self, name, a)
        return a

    # The underlying __enter__ method returns the wrong object
    # (self._fh) so override it to return the wrapper
    def __enter__(self):
        self._fh.__enter__()
        return self

class FilterInOutFiles(object):
    '''FilterInOutFiles: opens an input and an output file for a filter.
    
    Given an input file path, open that path for reading, 
    generate an output path, and open that path for writing.
    Make both path objects available as attributes.
    
    Intended for use with a command-line filter, which reads the
    contents of the input file and from this writes the output file.
    This class is independent of how the filter interprets the input
    data or generates the output data.

    >>> import os, os.path, tempfile
    >>> p = tempfile.mkdtemp()
    >>> f = open( os.path.join(p, 'test.txt'), 'w' )
    >>> f.write(''); f.close()
    >>> C = FilterInOutFiles('.out')
    >>> fh_i, fh_o = C.open_in_out_files(f.name)
    >>> os.path.basename(fh_o.name)
    'test.out.txt'
    >>> C.close()
    >>> os.remove(f.name); os.remove( fh_o.name );
    >>> os.rmdir(p)
    '''
    
    def __init__(self, output_ext='.out'):
        '''FilterInOutFiles(path): open path as an OFX file, prepare to repair'''
        
        self.output_ext = output_ext
        self.in_path = self.in_file = None
        self.out_path = self.out_file = None

    def generate_out_path(self, path):
        '''Generate an output file path based on given path.
        Path: Unicode string, path to input file.
        
        >>> C = FilterInOutFiles('.out')
        >>> C.generate_out_path('foo.txt')
        'foo.out.txt'
        '''

        if path is None or path == '' \
                    or self.output_ext is None or self.output_ext == '':
            return path
        
        (root, ext) = os.path.splitext(path)
        return root+self.output_ext+ext

    IN_FLAGS = 'r'  # flags to use with io.open() when opening in_path
    OUT_FLAGS = 'w' # flags to use with io.open() when opening out_path
    def open_in_out_files(self, in_path):
        '''open_in_out_files(in_path): return open inFile, outFile objects.
        '''
        self.in_path = in_path
        self.out_path = self.generate_out_path(in_path)
        self.in_file = open(self.in_path, self.IN_FLAGS)
        self.out_file = open(self.out_path, self.OUT_FLAGS)
        
        return (self.in_file, self.out_file)

    def close(self):
        '''close(): close the input and output files, erase the paths
        '''
        if self.in_file is not None:
            self.in_file.close()
        self.in_file = self.in_path = None
        if self.out_file is not None:
            self.out_file.close()
        self.out_file = self.out_path = None

    
class OFXRepairer(object):
    REPAIRED_EXT = '.repaired'  
    # repaired files have this extra extension before their extension
    # e.g. foo.ofx after repair is written to foo.repaired.ofx

    def __init__(self, path, repaired_ext=REPAIRED_EXT):
        '''OFXRepairer(path): open path as an OFX file, prepare to repair'''
        
        self.file_manager = FilterInOutFiles(repaired_ext)
        self.in_file = self.out_file = self._fh = self.codec_name = None
        if path is not None:
            self.in_file, self.out_file = \
                    self.file_manager.open_in_out_files(path)
            # OfxFile reads the headers, handles encoding. 
            # Its ._fh is a file object which decodes properly.
            f = OfxFile(self.in_file)
            self._fh = f._fh
            self.codec_name = self.codec_name_from_ofx_headers(f.headers)

    def __del__(self):
        '''OFXRepairer destructor: close file handles'''
        if self._fh is not None:
            self._fh.close()
        self.file_manager.close()

    def codec_name_from_ofx_headers(self, headers):
        '''From OFX headers dict, derive Python codec name
        
        headers: (ordered) dict, with ENCODING and CHARSET entries.
        '''
        # based on ofxparse.ofxparse.handle_encoding()
        enc = headers.get('ENCODING')

        if not enc:
            # no encoding specified, use 'ascii' codec
            return 'ascii'

        if enc == "USASCII":
            cp = headers.get("CHARSET", "1252")
            return "cp%s" % (cp, )

        elif enc in ("UNICODE", "UTF-8"):
            return "utf-8"
        
        raise AssertionError("OFX file lacks valid ENCODING and CHARSET entries.") 
        return None
        
        
    # Regular expression extracting content between OFX start and end elements
    RE_OFX = re.compile(r'(?is)([^<]*<OFX>)(.*?)(</OFX>.*)')
    
    def split_input(self, s):
        '''Split_input(s): split s into pre, to_repair, and post strings
        
        >>> r = OFXRepairer(None)
        >>> r.split_input("foo foo <OFX>stuff stuff stuff</OFX>bar bar")
        ('foo foo <OFX>', 'stuff stuff stuff', '</OFX>bar bar')
        '''
        m_ofx = self.RE_OFX.match(s)
        if m_ofx and m_ofx.lastindex == 3:
            # valid contents: return them
            return m_ofx.group(1), m_ofx.group(2), m_ofx.group(3)
        # Failed, return Nones
        return None, None, None

    # Regular expression extracting STMTTRN element
    RE_STMTTRN = re.compile(r'''(?isx)(?P<pre><STMTTRN>.*?\n)
                (?P<name_tag>\s*<NAME>)(?P<name_line>.*?)\n
                (?P<memo_tag>\s*<MEMO>)(?P<memo_line>.*?)
                    (?P<conf_field>(\s*Confirmation\s\#\d+\s*)?)\n
                (?P<post>.*?</STMTTRN>)''')
    
    # Expression to repair matches to RE_STMTTRN
    REPL = r'\g<pre>\g<name_tag>\g<memo_line>\n' \
           r'\g<memo_tag>\g<name_line>\g<conf_field>\n\g<post>'
    
    
    def repair(self, to_repair):
        '''repair(s): perform the repair on string s, returning repaired s
        
        >>> r = OFXRepairer(None)
        >>> print(r.repair("""\
<STMTTRN>\\n\
<DTPOSTED>20161201000000[-8:PST]\\n\
<NAME>Bill payment online\\n\
<MEMO>HYDRO 8509 Confirmation #743046       \\n\
<TRNAMT>-20.00\\n\
</STMTTRN>\
"""))
        <STMTTRN>
        <DTPOSTED>20161201000000[-8:PST]
        <NAME>HYDRO 8509
        <MEMO>Bill payment online Confirmation #743046       
        <TRNAMT>-20.00
        </STMTTRN>
                
        '''
        return self.RE_STMTTRN.sub(self.REPL, to_repair)
        # note: since re.sub() has no count, it substitutes all occurrences.
       
    def write(self):
        '''repair and write out the repaired file contents
        
        For this utility, we swap the values of the NAME and MEMO fields.
        '''
        
        with codecs.lookup(self.codec_name).streamwriter(self.out_file) as fh_out:
            # everything through the first <OFX> tag
            self._fh.seek(0)
            s = self._fh.read()
            pre, to_repair, post = self.split_input(s)
            if pre is None:
                raise CLIError('Appears to not be OFX: {0}'.format(self.in_path))
            else:                
                fh_out.write(pre)
                fh_out.write(self.do_repair(to_repair))
                fh_out.write(post)
        

def main(argv=None): # IGNORE:C0111
    '''Command line options.
    
    >>> print( "Exit Code: {0}".format(main(['foo.dat'])) )
    I don't work on files ending in '.dat': foo.dat.
    Exit Code: 0
    '''

    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_name = os.path.basename(sys.argv[0])
    program_version = "v%s" % __version__
    program_build_date = str(__updated__)
    program_version_message = '%%(prog)s %s (%s)' % (program_version, program_build_date)
    program_shortdesc = __import__('__main__').__doc__.split("\n")[1]
    program_license = '''%s

  Created by Jim DeLaHunt on %s.
  Main program in the public domain. Some modules copyright their authors,
  and released under the MIT licence.

  Distributed on an "AS IS" basis without warranties
  or conditions of any kind, either express or implied.

USAGE
''' % (program_shortdesc, str(__date__))

    try:
        # Setup argument parser
        parser = ArgumentParser(description=program_license, formatter_class=RawDescriptionHelpFormatter)
        # parser.add_argument("-r", "--recursive", dest="recurse", action="store_true", help="recurse into subfolders [default: %(default)s]")
        parser.add_argument("-v", "--verbose", dest="verbose", action="count", 
                            default=0, help="set verbosity level [default: %(default)s]")
        # parser.add_argument("-i", "--include", dest="include", help="only include paths matching this regex pattern. Note: exclude is given preference over include. [default: %(default)s]", metavar="RE" )
        # parser.add_argument("-e", "--exclude", dest="exclude", help="exclude paths matching this regex pattern. [default: %(default)s]", metavar="RE" )
        parser.add_argument('-V', '--version', action='version', version=program_version_message)
        parser.add_argument(dest="paths", help="paths to files(s) to repair [default: %(default)s]", 
                            metavar="path", nargs='+')

        # Process arguments
        args = parser.parse_args()

        paths = args.paths
        verbose = args.verbose
        # recurse = args.recurse
        # inpat = args.include
        # expat = args.exclude

        if verbose > 0:
            print("Verbose mode on")
#             if recurse:
#                 print("Recursive mode on")
#             else:
#                 print("Recursive mode off")
            print("Repairing {0} files: {1}".format(len(paths), paths))

        for inpath in paths:
            (_,ext) = os.path.splitext(inpath)
            if ext.lower() in ['.ofx', '.qfx']:
                if verbose > 0:
                    print("Repairing {0}...".format(inpath))
                r = OFXRepairer(inpath)
                r.write()
                print("Repaired {0}.".format(inpath))
            else:
                print("I don't work on files ending in '{0}': {1}.".format(ext, inpath))
        return 0
    except KeyboardInterrupt:
        ### handle keyboard interrupt ###
        return 0
    except Exception, e:
        if DEBUG or TESTRUN:
            raise(e)
        indent = len(program_name) * " "
        sys.stderr.write(program_name + ": " + repr(e) + "\n")
        sys.stderr.write(indent + "  for help use --help\n")
        return 2

if __name__ == "__main__":
    if DEBUG:
#        sys.argv.append("-h")
        sys.argv.append("-v")
#        sys.argv.append("-r")
    if TESTRUN:
        import doctest
        doctest.testmod()
    if PROFILE:
        import cProfile
        import pstats
        profile_filename = 'vanswap_ofx_profile.txt'
        cProfile.run('main()', profile_filename)
        statsfile = open("profile_stats.txt", "wb")
        p = pstats.Stats(profile_filename, stream=statsfile)
        stats = p.strip_dirs().sort_stats('cumulative')
        stats.print_stats()
        statsfile.close()
        sys.exit(0)
    sys.exit(main())