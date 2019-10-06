from cx_Freeze import setup, Executable

# executable options
script = 'principal.py'            # nome do arquivo principal .py a ser compilado
base = 'Win32GUI'                     # usar 'Win32GUI' para gui's e 'None' para console
#icon = 'icon_64.ico'              # nome (ou diretório) do icone do executável
targetName = 'Vendas_Mensais.exe'  # nome do .exe que será gerado
icon = None

# build options
packages = ['openpyxl', 'cx_Oracle', ]                 # lista de bibliotecas a serem incluídas
includes = []                                          # lista de módulos a serem incluídos
#include_files = ['icon_64.png', 'logo-90.png']        # lista de outros arquivos (imagens, dados...)
include_files = []

# shortcut options
shortcut_name = 'VdMsais'          # atalho criado no processo de instalação

# bdist_msi options
company_name = 'Dir_VendasMensais' # pasta em 'Arquivos de Programas' onde será instalado
product_name = 'Vendas_Mensais'    # subpasta em que o software será instalado
upgrade_code = '{66620F3A-DC3A-11E2-B341-002219E9B01E}' # código para upgrade de versão do programa
add_to_path = False                                     # adicionar o programa ao path?

# setup options
name = 'Relatorio de Vendas Mensais'    # Nome do programa na descrição
version = '1.2'                         # versão do programa na descrição
description = 'Sistema Para Lancamento de Relatorios Mensais (Balanco)'    # descrição do programa

"""
Edit the code above this comment.
Don't edit any of the code bellow.
"""

msi_data = {'Shortcut': [
    ("DesktopShortcut",                      # Shortcut
     "DesktopFolder",                        # Directory_
     shortcut_name,                          # Name
     "TARGETDIR",                            # Component_
     "[TARGETDIR]/{}".format(targetName),    # Target
     None,                                   # Arguments
     None,                                   # Description
     None,                                   # Hotkey
     None,                                   # Icon
     None,                                   # IconIndex
     None,                                   # ShowCmd
     "TARGETDIR",                            # WkDir
     ),

    ("ProgramMenuShortcut",                  # Shortcut
     "ProgramMenuFolder",                    # Directory_
     shortcut_name,                          # Name
     "TARGETDIR",                            # Component_
     "[TARGETDIR]/{}".format(targetName),    # Target
     None,                                   # Arguments
     None,                                   # Description
     None,                                   # Hotkey
     None,                                   # Icon
     None,                                   # IconIndex
     None,                                   # ShowCmd
     "TARGETDIR",                            # WkDir
     )
    ]
}

opt = {
    'build_exe': {'packages': packages,
                  'includes': includes,
                  'include_files': include_files
                  },
    'bdist_msi': {'upgrade_code': upgrade_code,
                  'add_to_path': add_to_path,
                  'initial_target_dir': r'[ProgramFilesFolder]\%s\%s' % (company_name, product_name),
                  'data': msi_data
                  }
}

exe = Executable(
    script=script,
    base=base,
    icon=icon,
    targetName=targetName,
    # shortcutName=shortcut_name,
    # shortcutDir='DesktopFolder'
)

setup(name=name,
      version=version,
      description=description,
      options=opt,
      executables=[exe]
      )

with open('version', 'w') as f:
	f.write('2.0.1')

