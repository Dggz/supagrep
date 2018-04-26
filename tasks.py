"""Management tasks"""

# pylint: disable=missing-param-doc, missing-type-doc, missing-docstring
import glob
import os
from itertools import chain

from invoke import task
from textwrap import dedent

CMD_EXE_PATH = os.environ['COMSPEC']


@task
def buildhash(ctx):
    """Create buildhash.py file with commit hash"""

    git_cmd = 'git rev-parse --short HEAD'
    commit_short_hash = ctx.run(git_cmd, shell=CMD_EXE_PATH).stdout.rstrip('\n')
    with open('supergrep/buildhash.py', mode='w') as f:
        f.write(dedent('''\
            """Build-time constants"""
            commit_short_hash = "{}"
        '''.format(commit_short_hash)))


@task
def topyx(ctx):
    """Rename all modules to cython modules"""

    for file in glob.glob('supergrep/**/*.py', recursive=True):
        if not file.endswith('__.py'):
            pyx_file = '{}.pyx'.format(os.path.splitext(file)[0])
            os.rename(file, pyx_file)


@task
def topy(ctx):
    """Rename back from cython modules to python ones"""

    for file in glob.glob('supergrep/**/*.pyx', recursive=True):
        py_file = '{}.py'.format(os.path.splitext(file)[0])
        os.rename(file, py_file)


@task
def cython(ctx):
    ctx.run('python setup.py build_ext --inplace', shell=CMD_EXE_PATH)


@task
def cleanup(ctx):
    """Clean after cython"""

    for file in chain.from_iterable(
            glob.glob(pattern, recursive=True) for pattern
            in ('supergrep/**/*.c', 'supergrep/**/*.pyd')):
        os.remove(file)


@task(pre=[buildhash, topyx, cython], post=[topy, cleanup])
def pyinstaller(ctx):
    """Call pyinstaller to create final exe"""

    ctx.run('pyinstaller --dist=build supergrep.spec', shell=CMD_EXE_PATH)
