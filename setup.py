import os
from distutils.core import setup


def get_version():
    with open(os.path.join('onepy', '__init__.py')) as f:
        for line in f:
            if line.startswith('__version__ ='):
                return line.split('=')[1].strip().strip('"\'')


setup(
    name = 'onepy',
    packages = ['onepy'],
    version = get_version(),
    description = 'OneNote Object Model',
    author='Varun Srinivasan',
    author_email='varunsrin@gmail.com',
    url="https://github.com/varunsrin/one-py",
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'Natural Language :: English',
        'License :: OSI Approved :: MIT License',
        'Environment :: Win32 (MS Windows)',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ], 
    test_suite="tests"
)