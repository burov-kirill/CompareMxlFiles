import pkgutil

from cx_Freeze import setup, Executable
executables = [Executable('main.py', base='Win32GUI',
                          target_name='CompareMXL.exe',
                          icon='ico/analysis_finance_statistics_business_graph_chart_report_icon_254045.ico')]
includefiles = ['__VERS__.txt']

def AllPackage():
    return [i.name for i in list(pkgutil.iter_modules()) if i.ispkg]
def notFound(A,v): # Check if v outside A
    try:
        A.index(v)
        return False
    except:
        return True
Import  = ['difflib', 'filecmp', 'glob', 'hashlib', 'os', 'queue', 're', 'sys', 'time', 'pandas',
        'PySimpleGUI', 'openpyxl', 'unicodedata','logging', 'subprocess', 'requests', 'win32com', 'numpy', 'pytz', 'dateutil', 'urllib', 'json',
           'ctypes', 'tkinter', 'email', 'http', 'xml', 'et_xmlfile', 'urllib3', 'charset_normalizer', 'idna', 'cetifi']

BasicPackages=["collections","encodings","importlib"] + Import
options = {
      'build_exe': {
          'include_files': includefiles,
            'build_exe': 'build_windows',
            # 'includes': BasicPackages,
            # 'excludes': [i for i in AllPackage() if notFound(BasicPackages,i)],
            "zip_include_packages": "*",
            "zip_exclude_packages": "",
            'optimize': 1
      }
}


setup(name='CompareMXL',
      version='1.0.1',
      description='Сверка MXL файлов',
      executables=executables,
      options=options)