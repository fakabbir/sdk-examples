# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

from com.sun.star.document import XEventListener
from com.sun.star.task import XJobExecutor
import os
import os.path
import pickle
import re
import shutil
import subprocess
import unittest
import tempfile
import threading
import time
import uno
import unohelper

class LawParagraph:
    '''This class represents a paragraph -- thats is a part of a law text'''
    def __init__(self, law, paragraph, startline, endline):
        (self.law, self.paragraph) = (law.lower(), paragraph.lower())
        (self.startline, self.endline) = (startline, endline)
    def get_name(self):
        return '%s %s' % (self.law, self.paragraph)

class ParagraphScanner:
    '''This class scans for the paragraphs in a lawtext and thus splits the text in such units'''
    paragraphre = re.compile('^#+ § ([0-9]+) .*')
    artikelre = re.compile('^#+ Art ([0-9]+).*')
    def create_paragraph_index(self, law, lawcontent):
        paragraphs = []
        current = None
        linecount = 0
        for line in lawcontent.split('\n'):
            paramatch = ParagraphScanner.paragraphre.match(line)
            artmatch = ParagraphScanner.artikelre.match(line)
            # TODO: Likely needs more regular expressions for some texts
            if paramatch or artmatch:
                if current:
                    current.endline = linecount
                    paragraphs.append(current)
            if paramatch:
                current = LawParagraph(law, paramatch.group(1), linecount, 0)
            elif artmatch:
                current = LawParagraph(law, artmatch.group(1), linecount, 0)
            linecount += 1
        if current:
            current.endline = linecount
            paragraphs.append(current)
        return paragraphs

class GitWrapper:
    '''This class wraps interaction with a git repository'''
    def __init__(self, repopath):
        self.repopath = repopath
    def get_filenames(self, branch):
        '''Returns a list of all the files that are checked in at one commit'''
        tree = subprocess.check_output(['git', '-C', self.repopath, 'ls-tree', '-r', '--name-only', '--full-tree', branch])
        return [ line for line in str(tree, encoding='utf8').split('\n') ]
    def get_filecontent(self, branch, filename):
        '''Returns the content of one file as UTF8 string'''
        return str(subprocess.check_output(['git', '-C', self.repopath, 'show', '%s:%s' % (branch, filename)], stderr=None), encoding='utf8', errors='ignore')
    def init_clone(self, fromrepo):
        '''Create a local clone from a remote repository'''
        subprocess.check_output(['git', 'clone', '--mirror', fromrepo, self.repopath], stderr=subprocess.DEVNULL)
    def update(self):
        '''Fetches all updates from the remote repository'''
        subprocess.check_output(['git', '-C', self.repopath, 'fetch', '--all'], stderr=subprocess.DEVNULL)
    def get_commithash(self, branch):
        '''Returns the commithash of the HEAD of a branch'''
        res = str(subprocess.check_output(['git', '-C', self.repopath, 'log', '-1', '--pretty=%H', branch], stderr=subprocess.DEVNULL), encoding='utf8').rstrip('\n')
        return res

class CommitIndex:
    '''This class holds an index of all the available paragraphs of all the laws at one commit'''
    def __init__(self, commithash):
        self.commithash = commithash
        self.paragraphs_by_name = {}
    def index(self, gitwrapper):
        '''Updates this index from the given git repo'''
        scanner = ParagraphScanner()
        (self.paragraphs_by_name, paragraphs) = ({}, [])
        for filename in gitwrapper.get_filenames(self.commithash):
            if os.path.basename(filename) == 'index.md':
                lawname = os.path.basename(os.path.dirname(filename))
                filecontent = gitwrapper.get_filecontent(self.commithash, filename)
                paragraphs.extend(scanner.create_paragraph_index(lawname, filecontent))
        for p in paragraphs:
            self.paragraphs_by_name[p.get_name()] = p
    def get_paragraph(self, gitwrapper, short_name):
        '''Extracts the content of one paragraph from the git repo''' 
        p = self.paragraphs_by_name[short_name.lower()]
        lawpath = os.path.join(p.law[0], p.law, 'index.md')
        lawtextlines = gitwrapper.get_filecontent(self.commithash, lawpath).split('\n')
        return '\n'.join(lawtextlines[p.startline:p.endline])
    def __get_state__(self):
        '''serializes this index into a state object (saving)'''
        state = [self.commithash]
        for p in self.paragraphs_by_name.values():
            state.append((p.law, p.paragraph, p.startline, p.endline))
        return state
    def __set_state__(self, state):
        '''deserializes this index from a state object (loading)'''
        self.commithash = state[0]
        for p_state in state[1:]:
            p = LawParagraph(p_state[0], p_state[1], p_state[2], p_state[3])
            self.paragraphs_by_name[p.get_name()] = p

# tests below this point

class TestParagraphScanner(unittest.TestCase):
    def test_default_roundtrip(self):
        scanner = ParagraphScanner()
        paras = []
        with open('./test-data/s/stgb/index.md', encoding='utf8') as stgbfile:
            paras = scanner.create_paragraph_index('stgb', stgbfile.read())
        self.assertGreater(len(paras), 0)
        with open('./test-data/g/gg/index.md', encoding='utf8') as ggfile:
            paras = scanner.create_paragraph_index('stgb', ggfile.read())
        self.assertGreater(len(paras), 0)

class TestGitWrapper(unittest.TestCase):
    def setUp(self):
        (self.testdir, self.clonedir) = (tempfile.mkdtemp(), None)
        subprocess.check_output(['git', 'init', self.testdir], stderr=subprocess.DEVNULL)
        subprocess.check_output(['git', '-C', self.testdir, 'commit', '--allow-empty', '-m', 'init'], stderr=subprocess.DEVNULL)
        with open(os.path.join(self.testdir, 'somefile'), 'w', encoding='utf8') as somefile:
            somefile.write('foo')
        with open(os.path.join(self.testdir, 'otherfile'), 'w', encoding='utf8') as otherfile:
            otherfile.write('bar\nbar')
        subprocess.check_output(['git', '-C', self.testdir, 'add', '-A'], stderr=None)
        subprocess.check_output(['git', '-C', self.testdir, 'commit', '-m', 'some files'], stderr=subprocess.DEVNULL)
    def tearDown(self):
        shutil.rmtree(self.testdir)
        if(self.clonedir):
            shutil.rmtree(self.clonedir)
    def test_get_filenames(self):
        gitwrapper = GitWrapper(self.testdir)
        filenames = gitwrapper.get_filenames('master')
        self.assertIn('somefile', filenames)
        self.assertIn('otherfile', filenames)
    def test_get_filecontent(self):
        gitwrapper = GitWrapper(self.testdir)
        content = gitwrapper.get_filecontent('master', 'somefile')
        self.assertEqual(content, 'foo')
        content = gitwrapper.get_filecontent('master', 'otherfile')
        self.assertEqual(content, 'bar\nbar')
    def test_init_clone(self):
        self.clonedir = tempfile.mkdtemp()
        gitwrapper = GitWrapper(self.clonedir)
        gitwrapper.init_clone(self.testdir)
        filenames = gitwrapper.get_filenames('master')
        self.assertIn('somefile', filenames)
        self.assertIn('otherfile', filenames)
    def test_update(self):
        self.clonedir = tempfile.mkdtemp()
        gitwrapper = GitWrapper(self.clonedir)
        gitwrapper.init_clone(self.testdir)
        with open(os.path.join(self.testdir, 'thirdfile'), 'w', encoding='utf8') as thirdfile:
            thirdfile.write('baz')
        subprocess.check_output(['git', '-C', self.testdir, 'add', '-A'], stderr=None)
        subprocess.check_output(['git', '-C', self.testdir, 'commit', '-m', 'third file'], stderr=subprocess.DEVNULL)
        gitwrapper.update()
        filenames = gitwrapper.get_filenames('master')
        self.assertIn('thirdfile', filenames)

if __name__ == '__main__':
    unittest.main()

# UNO plumbing here

class SyncRepoThread(threading.Thread):
    def run(self):
        '''This thread downloads, updates and reindexes the BundesGit repo'''
        gitwrapper = GitWrapper(self.repopath)
        if not os.path.isdir(self.repopath):
            os.mkdir(self.repopath)
            gitwrapper.init_clone('https://github.com/bundestag/gesetze.git')
        gitwrapper.update()
        commitindex = CommitIndex('')
        if os.path.isfile(self.indexpath):
            with open(self.indexpath, 'rb') as indexfile:
                commitindex.__set_state__(pickle.load(indexfile))
        newhash = gitwrapper.get_commithash('master')
        if not commitindex.commithash == newhash:
            commitindex.commithash = newhash
            commitindex.index(gitwrapper)
            with open(self.indexpath, 'wb') as indexfile:
                pickle.dump(commitindex.__get_state__(), indexfile)

class BundesGit(unohelper.Base, XJobExecutor, XEventListener):
    '''Class that implements the service registered in LibreOffice reacting to events there:

        see http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1task_1_1JobExecutor.html
        see http://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1document_1_1XEventListener.html
    '''
    def __md_to_text(self, mdtext):
        '''Primitive markdown to LibreOffice text translation'''
        # TODO: could use more sofistication, e.g. translate enumerations etc.
        result = ''
        newline = True
        for line in mdtext.split('\n'):
            if len(line):
                if not newline:
                    result = result + ' '
                result = result + line
                newline = False
            else:
                result = result + '\n'
                newline = True
        return result
    def trigger(self, args):
        if args == 'onFirstVisibleTask':
            # LibreOffice started: download or update our git repo in the background
            syncrepo = SyncRepoThread()
            syncrepo.repopath = self.__get_repopath()
            syncrepo.indexpath = self.__get_indexpath()
            syncrepo.start()
        elif args == 'insertlaw':
            # inserting a law was requested
            gitwrapper = GitWrapper(self.__get_repopath())
            # load the latest commitindex
            commitindex = CommitIndex('')
            with open(self.__get_indexpath(), 'rb') as indexfile:
                commitindex.__set_state__(pickle.load(indexfile))
            # select the current paragraph
            desktop = self.createUnoService("com.sun.star.frame.Desktop")
            component = desktop.CurrentComponent
            controller = component.CurrentController
            cursor = controller.ViewCursor
            cursor.gotoStartOfLine(False)
            cursor.gotoEndOfLine(True)
            # get the text in the current paragraph
            lawtext = commitindex.get_paragraph(gitwrapper, cursor.String)
            if lawtext:
                # if that references an abbrevation for a paragraph, replace the text
                cursor.String = self.__md_to_text(lawtext)
            else:
                # otherwise, just move to the end of the paragraph
                cursor.collapseToEnd()
    # boilerplate code below this point
    def __init__(self, context):
        self.context = context
        self.workdir = None
    def __find_workdir(self):
        '''Find the place where the extension has been installed to'''
        return os.path.dirname(str(__file__, encoding='utf8'))
    def __get_repopath(self):
        return os.path.join(self.__find_workdir(), 'bundesgitrepo')
    def __get_indexpath(self):
        return os.path.join(self.__find_workdir(), 'commitindex')
    def createUnoService(self, name):
        return self.context.ServiceManager.createInstanceWithContext(name, self.context)
    def disposing(self, args):
        pass
    def notifyEvent(self, args):
        pass

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    BundesGit,
    'org.libreoffice.bundesgit.BundesGit',
    ('com.sun.star.task.JobExecutor',))
