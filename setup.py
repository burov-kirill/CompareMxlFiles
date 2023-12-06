from cx_Freeze import setup, Executable
executables = [Executable('main.py', base='Win32GUI',
                          target_name='CompareMXL.exe',
                          icon='ico/analysis_finance_statistics_business_graph_chart_report_icon_254045.ico')]
includefiles = ['__VERS__.txt']



options = {
      'build_exe': {
          'include_files': includefiles,
            'build_exe': 'build_windows',
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