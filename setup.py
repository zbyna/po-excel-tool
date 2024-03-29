from setuptools import setup, find_packages
import sys

version = '0.1.0'

install_requires=[
        'click',
        'polib',
        'openpyxl',
        'argparse;python_version<"3.0"',
        'tqdm',
        ]

setup(name='poexceltool',
      version=version,
      description='Convert between Excel and PO files',
      long_description=open('README.md').read() + '\n' + \
              open('changes.md').read(),
      long_description_content_type='text/markdown',
      classifiers=[
          'Environment :: Console',
          'Intended Audience :: Developers',
          'License :: DFSG approved',
          'License :: OSI Approved :: BSD License',
          'Operating System :: OS Independent',
          'Programming Language :: Python :: 2',
          'Programming Language :: Python :: 2.6',
          'Programming Language :: Python :: 2.7',
          'Programming Language :: Python :: 3',
          'Intended Audience :: Developers',
          'Intended Audience :: End Users/Desktop',
          ],
      keywords='translation po gettext Babel lingua',
      author='Zbyněk Fiala',
      author_email='zbynek.fiala@gmail.com',
      url='https://github.com/zbyna/po-excel-tool',
      license='BSD',
      packages=find_packages('src'),
      package_dir={'': 'src'},
      include_package_data=True,
      zip_safe=True,
      install_requires=install_requires,
      entry_points='''
      [console_scripts]
      pet = poexceltool.poexcel:poexcel
      '''
      )
