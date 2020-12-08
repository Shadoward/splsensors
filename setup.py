# -*- coding: utf-8 -*-
# Learn more: https://github.com/kennethreitz/setup.py

from setuptools import setup, find_namespace_packages

with open('README.md') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

install_requires = [
    'pandas',
    'glob3',
    'tqdm',
    'gooey',
    'obspy',
    'tabulate',
]

extras_require = {
    'build' : [
        'setuptools',
    ],
    'tests' : [],
}

setup(
    name='splsensors',
    version='0.3.4',
    description='Linename comparison and rename tool between SPL and sensors',
    long_description=readme,
    author='Patrice Ponchant',
    author_email='patrice.ponchant@fugro.com',
    include_package_data = True,
    install_requires=install_requires,
    extras_require=extras_require,
    tests_require=extras_require['tests'],
    url='',
    license=license,
    packages=find_namespace_packages(where='src'),
    package_dir={'': 'src'},
    keywords='FBF FBZ POS MBES SSS SBP MAG Comparison Rename',
    classifiers=[
        'Development Status :: 4 - Beta',
        'Natural Language :: English',
        'Topic :: Scientific/Engineering'
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)