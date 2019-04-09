# -*- coding: utf-8 -*-
"""
# Copyright (c) 2014 Patricio Moracho <pmoracho@gmail.com>
#
# db2xlsx.
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.    See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
"""
__author__        = "Patricio Moracho <pmoracho@gmal.com>"
__appname__        = "autoxls"
__appdesc__        = "Generación automatizada de archivos Excel"
__license__        = 'GPL v3'
__copyright__    = "2019 %s" % (__author__)
__version__        = "0.9"
__date__        = "2019/04/08"

"""
###############################################################################
# Imports
###############################################################################
"""
try:

    import sys
    import gettext
    from gettext import gettext as _
    gettext.textdomain('padrondl')

    def my_gettext(s):
        """my_gettext: Traducir algunas cadenas de argparse."""
        current_dict = {
            'usage: ': 'uso: ',
            'optional arguments': 'argumentos opcionales',
            'show this help message and exit': 'mostrar esta ayuda y salir',
            'positional arguments': 'argumentos posicionales',
            'the following arguments are required: %s': 'los siguientes argumentos son requeridos: %s',
            'show program''s version number and exit': 'Mostrar la versión del programa y salir',
            'expected one argument': 'se espera un valor para el parámetro',
            'expected at least one argument': 'se espera al menos un valor para el parámetro'
        }

        if s in current_dict:
            return current_dict[s]
        return s

    gettext.gettext = my_gettext

    import argparse
    import os
    import logging
    import traceback
    from xlsxWriter import xlsxWriter

except ImportError as err:
    modulename = err.args[0].split()[3]
    print("No fue posible importar el modulo: %s" % modulename)
    sys.exit(-1)


def delete_file(filename):
    try:
        os.remove(filename)
    except OSError:
        pass


def init_argparse():
    """Inicializar parametros del programa."""
    cmdparser = argparse.ArgumentParser(prog=__appname__,
                                        description        = "%s (v%s)\n%s\n" % (__appdesc__,__version__,__copyright__ ),
                                        epilog="",
                                        add_help=True,
                                        formatter_class=lambda prog: argparse.HelpFormatter(prog, max_help_position=45)
    )

    opciones = {    "yamlconfigfile": {
                                "type": str,
                                "nargs": '?',
                                "action": "store",
                                "help": _("Archivo de entrada (YAML)"),
                                "metavar": "\"archivo.yml\""
                    },
                    "--version -v": {
                                "action":    "version",
                                "version":    __version__,
                                "help":        _("Mostrar el número de versión y salir")
                    },
                    "--outputpath -o": {
                                "type":        str,
                                "action":    "store",
                                "dest":        "outputpath",
                                "default":    ".",
                                "help":        _("Carpeta de salida dónde se almacenaran las planillas"),
                                "metavar": "\"path\""
                    },
                    "--log-level -n": {
                                "type":        str,
                                "action":    "store",
                                "dest":        "loglevel",
                                "default":    "info",
                                "help":        _("Nivel de log")
                    },
                    "--log-file -l": {
                                "type":        str,
                                "action":    "store",
                                "dest":        "logfile",
                                "default":    None,
                                "help":        _("Archivo de log"),
                                "metavar":    "file"
                    }, 
                    "--keywordsfile -f": {
                                "type":        str,
                                "action":    "store",
                                "dest":        "keywordsfile",
                                "default":    None,
                                "help":        _("Archivo de keyword del proceso"),
                                "metavar":    """archivo"""
                    },
                    "--keywords -k": {
                                "type":        str,
                                "action":    "store",
                                "dest":        "keywords",
                                "default":    None,
                                "help":        _("Keywords del proceso de conversión"),
                                "metavar":    """{'key':'value','key':'value'}"""
                    },


                }


    for key, val in opciones.items():
        args = key.split()
        kwargs = {}
        kwargs.update(val)
        cmdparser.add_argument(*args, **kwargs)

    return cmdparser

def file_accessible(filepath, mode):
    """Check if a file exists and is accessible. """
    try:
        with open(filepath, mode, encoding='utf8'):
            pass
    except IOError:
        return False

    return True

"""
##################################################################################################################################################
# Main program
##################################################################################################################################################
"""
if __name__ == "__main__":

    # determine if application is a script file or frozen exe
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)

    # Para el proceso autmático/defaul (cuando invocamos sin parámetros)
    # esperamos que exista el archivo de configuración y de keywords
    default_yaml_file = os.path.join(application_path, 'db2xlsx.yaml')
    default_keywords_file = os.path.join(application_path, 'db2xlsx.keywords.yaml')

    cmdparser = init_argparse()
    try:
        args = cmdparser.parse_args()
    except IOError as msg:
        cmdparser.error(str(msg))
        sys.exit(-1)

    if not args.yamlconfigfile:
        # Si no se pasa el archivo, y existe 
        # el db2xlsx.yaml se trata de una ejecución automática
        if file_accessible(default_yaml_file, 'r'):
            args.default_yaml_file = default_yaml_file
            args.startexcel = True
            args.logfile = 'db2xlsx.log'
        else:
            cmdparser.error(u"debe indicar el archivo de input (--yamlconfigfile)")
            sys.exit(-1)

    log_level = getattr(logging, args.loglevel.upper(), None)
    logging.basicConfig(filename=args.logfile, level=log_level, format='%(asctime)s|%(name)s|%(levelname)s|%(message)s', datefmt='%Y/%m/%d %I:%M:%S', filemode='w')

    if args.outputpath == '{desktop}':
        outputpath = os.path.join(os.path.expanduser('~'), 'Desktop')
    else:
        if args.outputpath == '{tmp}':
            outputpath = tempfile._get_default_tempdir()
        else:
            outputpath = args.outputpath

    try:
        xlsw = xlsxWriter()
        xlsw.setup_writer_from_file(args.yamlconfigfile)
        xlsw.setup_logging_object(logging)
        xlsw.setup_outputpath(outputpath)

        if args.keywords:
            xlsw.add_keywords_from_string(args.keywords)
        else:
            if not args.keywordsfile:
                if file_accessible(default_keywords_file, 'r'):
                    xlsw.add_keywords_from_yamlfile(default_keywords_file)
            else:
                xlsw.add_keywords_from_yamlfile(args.keywordsfile)

    except Exception as e:
        sys.exit(-1)

    xlsw.process()
    
    """

    engine = Engine(jsonfile, keywords, logging)

    try:
        engine.generate(outputpath, args.startexcel)
        if args.dropcfgfiles:
            delete_file(jsonfile)
            delete_file(args.keyworfilejson)

    except Exception as e:
        logging.error("%s error: %s" % (__appname__, str(e)))

    """

    sys.exit(0)
