#!/usr/bin/env python2.7
# encoding: utf-8
'''
vanswap_ofx -- swap NAME and MEMO fields in OFX files 

It defines a class OFXRepairer, which does the work of parsing the
OFX file, performing the repair, and writing out the file with a 
modified filename.

@author:     Jim DeLaHunt

@copyright:  Main program is granted to the public domain. Some modules are copyright by their authors.

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
__version__ = 0.5
__date__ = '2017-03-25'
__updated__ = __date__

DEBUG = 0
TESTRUN = 0
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
    
    If there is already a file at the output path, raise an OSError 
    exception, with errno.EEXIST .
    >>> fh_i, fh_o = C.open_in_out_files(f.name)     # doctest: +IGNORE_EXCEPTION_DETAIL
    Traceback (most recent call last):
      ...
    OSError: [Errno 17] File exists: ...
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

    IN_FLAGS = 'rb'  # flags to use with open() when opening in_path
    OUT_O_FLAGS = os.O_WRONLY | os.O_CREAT | os.O_EXCL # flags to use with os.open() when opening out_path
    OUT_FLAGS = 'wb' # flags to use with open() when opening out_path
    def open_in_out_files(self, in_path):
        '''open_in_out_files(in_path): return open inFile, outFile objects.
        '''
        self.in_path = in_path
        self.out_path = self.generate_out_path(in_path)
        self.in_file = open(self.in_path, self.IN_FLAGS)
        # Crude check to prevent overwriting. os.open(path, os.O_CREAT | os.O_EXCL)
        # is a more reliable way, but leaves out_file.name not set to the path.
        if os.path.exists(self.out_path):
            import errno
            raise OSError(errno.EEXIST, 'File exists', self.out_path)
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
    def __init__(self, in_file=None, out_file=None):
        r'''OFXRepairer(in_file, out_file): prepare to repair OFX
        
        Instantiate with file objects for input and output to perform
        a repair. Caller must open and close file objects.
        
        You can instantiate without file parameters in test fixtures,
        in order to exercise the methods. 
        
        # Test a complete file example
        >>> import io, tempfile, os
        >>> in_file = io.BytesIO( u"""OFXHEADER:100
        ... DATA:OFXSGML
        ... VERSION:102
        ... SECURITY:TYPE1
        ... ENCODING:USASCII
        ... CHARSET:1252
        ... 
        ... <OFX>
        ...  <BANKMSGSRSV1>
        ...   <STMTTRNRS>
        ...    <STMTRS>
        ...      <STMTTRN>
        ...       <TRNTYPE>DEBIT
        ...       <DTPOSTED>20170101000000[-8:PST]
        ...       <TRNAMT>-12.34
        ...       <FITID>25.030001    1790116941000
        ...       <NAME>payment
        ...       <MEMO>VISA Confirmation #881665       
        ...      </STMTTRN>
        ...    </STMTRS>
        ...   </STMTTRNRS>
        ...  </BANKMSGSRSV1>
        ... </OFX>
        ... """.encode('cp1252'))
        >>> try:
        ...     (out_fd, out_path) = tempfile.mkstemp(); out_file = os.fdopen(out_fd, 'w')
        ...     r = OFXRepairer(in_file, out_file); r.write()
        ...     out_file = io.open(out_path, 'rb'); bb = out_file.readlines();
        ... finally:
        ...     os.remove(out_path)
        >>> bb[16].strip()
        '<NAME>VISA'
        >>> bb[17].strip()
        '<MEMO>payment Confirmation #881665'
        '''
        
        self.out_file = out_file
        self.in_file = self.codec_name = None
        if in_file is not None:
            # OfxFile reads the headers, and it handles encoding. 
            # Its ._fh is a file object which decodes properly.
            f = OfxFile(in_file)
            self.in_file = f.fh
            self.in_file.seek(0)  # Later, we reread from beginning
            self.codec_name = self.codec_name_from_ofx_headers(f.headers)

    def __del__(self):
        '''OFXRepairer destructor: close file handles'''
        if self.in_file is not None:
            self.in_file.close()

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
        r'''Split_input(s): split string s into pre, to_repair, and post strings

        RE_OFX is a regular expression which recognises three groups:
        pre-target, target, and post-target.  
        An <OFX> element ends the pre-target. An </OFX> starts the post-target.
        Given a string that matches this RE, return the three parts.
        >>> r = OFXRepairer(None)
        >>> r.split_input("foo foo <OFX>stuff stuff stuff</OFX>bar bar")
        ('foo foo <OFX>', 'stuff stuff stuff', '</OFX>bar bar')
        
        Given a string that does not match this RE, return three Nones.
        >>> r.split_input("Has <OFX> element, but does not fully match.")
        (None, None, None)
        >>> r.split_input("Does not match at all.")
        (None, None, None)
        
        A full-file example:
        >>> (pre, to_repair, post) = r.split_input("""OFXHEADER:100
        ... DATA:OFXSGML
        ... VERSION:102
        ... SECURITY:TYPE1
        ... ENCODING:USASCII
        ... CHARSET:1252
        ... 
        ... <OFX>
        ...  <BANKMSGSRSV1>
        ...   <STMTTRNRS>
        ...    <STMTRS>
        ...      <STMTTRN>
        ...       <TRNTYPE>DEBIT
        ...       <DTPOSTED>20170101000000[-8:PST]
        ...       <TRNAMT>-12.34
        ...       <FITID>25.030001    1790116941000
        ...       <NAME>payment
        ...       <MEMO>VISA Confirmation #881665       
        ...      </STMTTRN>
        ...    </STMTRS>
        ...   </STMTTRNRS>
        ...  </BANKMSGSRSV1>
        ... </OFX>
        ... """)
        >>> print(pre)
        OFXHEADER:100
        DATA:OFXSGML
        VERSION:102
        SECURITY:TYPE1
        ENCODING:USASCII
        CHARSET:1252
        <BLANKLINE>
        <OFX>
        >>> print(to_repair)
        <BLANKLINE>
         <BANKMSGSRSV1>
          <STMTTRNRS>
           <STMTRS>
             <STMTTRN>
              <TRNTYPE>DEBIT
              <DTPOSTED>20170101000000[-8:PST]
              <TRNAMT>-12.34
              <FITID>25.030001    1790116941000
              <NAME>payment
              <MEMO>VISA Confirmation #881665       
             </STMTTRN>
           </STMTRS>
          </STMTTRNRS>
         </BANKMSGSRSV1>
        <BLANKLINE>
        >>> print(post)
        </OFX>
        <BLANKLINE>

        '''

        m_ofx = self.RE_OFX.match(s)
        if m_ofx and m_ofx.lastindex == 3:
            # valid contents: return them
            return m_ofx.group(1), m_ofx.group(2), m_ofx.group(3)
        # Failed, return Nones
        return None, None, None


    # Regular expression extracting STMTTRN element
    RE_STMTTRN = re.compile(r'''(?ix)
                (?P<pre><STMTTRN>\s*\n(.*\n)*?)
                # Require both <NAME> and <MEMO> if we are to repair
                (?P<name_tag>\s*<NAME>)(?P<name_line>.*?)\n
                (?P<memo_tag>\s*<MEMO>)(?P<memo_line>.*?)
                      (?P<conf_field>(\s*Confirmation\s\#\d+\s*)?)\n
                (?P<post>(.*?\n)*?\s*</STMTTRN>)''')
    
    def repair(self, to_repair):
        r'''repair(to_repair): perform the repair on string s, returning repaired s
        
        Perform a repair on string to_repair. 
        The incorrect files have text after the <NAME> which belongs in
        the <MEMO>. The text in <MEMO> has a leading part which belongs
        in the <NAME>. However it may also have a trailing part of the
        form "Confirmation #1234". This should remain in the <MEMO>.
        
        # This is the desired repair.
        >>> r = OFXRepairer(None)
        >>> print(r.repair("""<STMTTRN>
        ... <DTPOSTED>20161201000000[-8:PST]
        ... <NAME>Bill payment online
        ... <MEMO>HYDRO 8509 Confirmation #743046       
        ... <TRNAMT>-20.00
        ... </STMTTRN>"""))
        <STMTTRN>
        <DTPOSTED>20161201000000[-8:PST]
        <NAME>HYDRO 8509
        <MEMO>Bill payment online Confirmation #743046       
        <TRNAMT>-20.00
        </STMTTRN>
                
        # Should also work if <MEMO> is the final element in <STMTTRN>
        >>> r = OFXRepairer(None)
        >>> print(r.repair("""<STMTTRN>
        ... <DTPOSTED>20161205000000[-8:PST]
        ... <TRNAMT>-20.00
        ... <NAME>Bill payment online
        ... <MEMO>HYDRO 8509 Confirmation #743046       
        ... </STMTTRN>"""))
        <STMTTRN>
        <DTPOSTED>20161205000000[-8:PST]
        <TRNAMT>-20.00
        <NAME>HYDRO 8509
        <MEMO>Bill payment online Confirmation #743046       
        </STMTTRN>
                
        Repair of a transaction with no Confirmation part.
        >>> print(r.repair("""<STMTTRN>
        ... <DTPOSTED>20161202000000[-8:PST]
        ... <NAME>Bill payment online 
        ... <MEMO>HYDRO 8511       
        ... <TRNAMT>-20.00
        ... </STMTTRN>"""))
        <STMTTRN>
        <DTPOSTED>20161202000000[-8:PST]
        <NAME>HYDRO 8511       
        <MEMO>Bill payment online 
        <TRNAMT>-20.00
        </STMTTRN>
                
        Some transactions have only a <NAME>, no <MEMO>. Those are unchanged.
        >>> print(r.repair("""<STMTTRN>
        ...   <TRNTYPE>CREDIT
        ...   <DTPOSTED>20161230100000[-8:PST]
        ...   <TRNAMT>1.02
        ...   <NAME>Interest credited to account
        ... </STMTTRN>"""))
        <STMTTRN>
          <TRNTYPE>CREDIT
          <DTPOSTED>20161230100000[-8:PST]
          <TRNAMT>1.02
          <NAME>Interest credited to account
        </STMTTRN>

        Transactions missing either <NAME> and <MEMO> field are not changed.
        >>> print(r.repair("""<STMTTRN>
        ... <TRNAMT>-10.00
        ... <DTPOSTED>20161203000000[-8:PST]
        ... </STMTTRN>"""))
        <STMTTRN>
        <TRNAMT>-10.00
        <DTPOSTED>20161203000000[-8:PST]
        </STMTTRN>

        This case confused the code: a transaction with only a <NAME>
        element, followed by one with <NAME> and <MEMO>. The first 
        must not change, the second changes only its own <NAME> element,
        it doesn't reach back to the first transaction's <NAME>.
        >>> print(r.repair("""<STMTTRN>
        ...   <NAME>Interest credited to account
        ... </STMTTRN>
        ... <STMTTRN>
        ...   <NAME>Funds transfer online
        ...   <MEMO>from Pay As You Go Chequing
        ... </STMTTRN>"""))
        <STMTTRN>
          <NAME>Interest credited to account
        </STMTTRN>
        <STMTTRN>
          <NAME>from Pay As You Go Chequing
          <MEMO>Funds transfer online
        </STMTTRN>
        
        This case confused the code: whitespace after opening tag 
        or before closing tag:
        >>> print(r.repair("""     <STMTTRN>  
        ...       <NAME>payment
        ...       <MEMO>VISA Confirmation #881665       
        ...      </STMTTRN>"""))
             <STMTTRN>  
              <NAME>VISA
              <MEMO>payment Confirmation #881665       
             </STMTTRN>

        The transaction part of the full example from __init__:
        >>> print(r.repair("""
        ...  <BANKMSGSRSV1>
        ...   <STMTTRNRS>
        ...    <STMTRS>
        ...      <STMTTRN>
        ...       <TRNTYPE>DEBIT
        ...       <DTPOSTED>20170101000000[-8:PST]
        ...       <TRNAMT>-12.34
        ...       <FITID>25.030001    1790116941000
        ...       <NAME>payment
        ...       <MEMO>VISA Confirmation #881665       
        ...      </STMTTRN>
        ...    </STMTRS>
        ...   </STMTTRNRS>
        ...  </BANKMSGSRSV1>
        ... """))
        <BLANKLINE>
         <BANKMSGSRSV1>
          <STMTTRNRS>
           <STMTRS>
             <STMTTRN>
              <TRNTYPE>DEBIT
              <DTPOSTED>20170101000000[-8:PST]
              <TRNAMT>-12.34
              <FITID>25.030001    1790116941000
              <NAME>VISA
              <MEMO>payment Confirmation #881665       
             </STMTTRN>
           </STMTRS>
          </STMTTRNRS>
         </BANKMSGSRSV1>
        <BLANKLINE>
        
        End of tests.
        '''

        def repl(m):
            """Safely generate a replace string from a match object
            
            We can't use re.sub() with named groups, because Python before 3.4
            throws "unmatched group" errors instead of substituting ''.
            """
            g = m.groupdict(default='')
            new_name = g['name_tag']+g['memo_line']+'\n' \
                            if g['name_tag'] else ''
            new_memo = g['memo_tag']+g['name_line']+g['conf_field']+'\n' \
                            if g['memo_tag'] else ''
            return g['pre'] + new_name + new_memo + g['post']

        return self.RE_STMTTRN.sub(repl, to_repair)
        # note: since re.sub() has no count, it substitutes all occurrences.
       

    def write(self):
        '''repair and write out the repaired file contents
        
        '''
        if self.out_file is None:
            return
        
        with codecs.lookup(self.codec_name).streamwriter(self.out_file) as fh_out:
            s = self.in_file.read()
            pre, to_repair, post = self.split_input(s)
            if pre is None:
                raise CLIError('Appears to not be OFX: {0}'.format(self.in_file.name))
            else:                
                fh_out.write(pre)
                fh_out.write(self.repair(to_repair))
                fh_out.write(post)
        

def main(argv=None): # IGNORE:C0111
    '''Command line options.

    Only works on files with specific extensions.
    However, the exit code is 0, not an error exit code.
    >>> sys.argv[1:] = ['foo.dat']
    >>> main()
    vanswap_ofx.py: vanswap_ofx -- swap NAME and MEMO fields in OFX files 
    <BLANKLINE>
    I don't work on files ending in '.dat': foo.dat.
    0
    
    If the input file doesn't exist, it prints an error message and continues.
    >>> import os, os.path, tempfile
    >>> p = tempfile.mkdtemp()
    >>> sys.argv[1:] = [ os.path.join(p, 'nonexistent.ofx') ]
    >>> main()        # doctest: +ELLIPSIS
    vanswap_ofx.py: vanswap_ofx -- swap NAME and MEMO fields in OFX files 
    <BLANKLINE>
    SORRY: File '...nonexistent.ofx' doesn't appear to exist.
    0

    If the output file exists, it prints an error message and continues.
    >>> f1 = open( os.path.join(p, 'existing.ofx'), 'w' ); f1.write(''); f1.close()
    >>> f2 = open( os.path.join(p, 'existing.repaired.ofx'), 'w' ); f2.write(''); f2.close()
    >>> sys.argv[1:] = [ f1.name ]
    >>> main()        # doctest: +ELLIPSIS
    vanswap_ofx.py: vanswap_ofx -- swap NAME and MEMO fields in OFX files 
    <BLANKLINE>
    SORRY: Output file '...existing.repaired.ofx' already exists, so unable to repair '...existing.ofx'.
    0

    >>> os.remove(f1.name); os.remove( f2.name );
    >>> os.rmdir(p)
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

vanswap_ofx is a utility to repair OFX files by swapping the NAME 
and MEMO fields. It works around a problem with OFX files created 
by my credit union after a system upgrade in 2016. They generated OFX
files with the NAME string in the MEMO field, and vice versa. This 
utility repairs the OFX file by reading it in, and switching the values
of NAME and MEMO in each transaction. If there is a confirmation number 
in the MEMO field, it stays in the MEMO field. Then it writes the 
repaired content out to a sister file, with a ".repaired" subextension, 
in the same directory. 

e.g. the command: vanswap.py statements/transactions_201610.ofx
writes repaired content to   statements/transactions_201610.repaired.ofx

Updated by Jim DeLaHunt on %s.
Main program is granted to the public domain. Some modules are copyright 
by their authors, and released under the MIT licence.
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
        
        print("{0}: {1}\n".format(program_name, program_shortdesc))

        if verbose > 0:
            print("Verbose mode on")
#             if recurse:
#                 print("Recursive mode on")
#             else:
#                 print("Recursive mode off")
            print("Repairing {0} files: {1}".format(len(paths), paths))

        file_manager = FilterInOutFiles('.repaired')
        # repaired files have this extra extension before their extension
        # e.g. foo.ofx after repair is written to foo.repaired.ofx

        for inpath in paths:
            (_,ext) = os.path.splitext(inpath)
            if ext.lower() in ['.ofx', '.qfx']:
                if verbose > 0:
                    print("Repairing {0}...".format(inpath))
                try:
                    in_file, out_file = file_manager.open_in_out_files(inpath)
                except (IOError, OSError), e:
                    import errno
                    if e.errno == errno.ENOENT:
                        print("SORRY: File '{0}' doesn't appear to exist.".format(e.filename))                        
                    elif e.errno == errno.EEXIST:
                        print("SORRY: Output file '{1}' already exists, so unable to repair '{0}'.".format(inpath, e.filename))
                    else:
                        print("SORRY: Unable to repair '{0}', because exception '{1}' occurred.".format(inpath, e))
                    break # give up in inpath, go on to next
                
                r = OFXRepairer(in_file, out_file)
                r.write()
                print("Copy of '{0}' repaired, in '{1}'.".format(inpath, out_file.name))
            else:
                print("I don't work on files ending in '{0}': {1}.".format(ext, inpath))
            file_manager.close()

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
        sys.exit(0)
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