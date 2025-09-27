import PyInstaller.__main__
import shutil
import os

app_path = str(os.path.dirname(os.path.realpath(__file__)) + '\\app.py')
assets_path = str(os.path.dirname(os.path.realpath(__file__)) + '\\assets')
icon_path = str(os.path.dirname(os.path.realpath(__file__)) + '\\assets\\icon.ico')
build_path = str(os.path.dirname(os.path.realpath(__file__)) + '\\dist\\assets')
version_path = str(os.path.dirname(os.path.realpath(__file__)) + '\\version.txt')

def build():
 shutil.rmtree('dist', ignore_errors=True)
 shutil.rmtree('build', ignore_errors=True)

 PyInstaller.__main__.run([
  app_path,
  '--name=ALR',
  '--onefile',
  '--noconsole',
  '--windowed',
  '--icon=' + icon_path,
  '--version-file=' + version_path,
  '--add-data=' + icon_path + ';.',
 ])

if __name__ == '__main__':
 build()
 print("Build completed successfully.")