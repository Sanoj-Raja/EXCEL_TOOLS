from setuptools import setup

classifiers = [
  'Development Status :: 4 - Beta',
  'Intended Audience :: Education',
  'License :: OSI Approved :: MIT License',
  'Operating System :: OS Independent',
  'Programming Language :: Python :: 3.9'
]
 
setup(
  name='excel_tools', 
  version='1.0',
  description='A custom Python Package to automate the manual copy paste excel work.',
  long_description=open('README.md').read(), 
  # Always keep the readme file in capital letters. 
  author='Sanoj Raja',
  classifiers=classifiers,
  keywords='excel tools', 
  packages= ['excel_tools'], 
  install_requires=['openpyxl'] 
)