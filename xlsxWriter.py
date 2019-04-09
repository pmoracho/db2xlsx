# -*- coding: utf-8 -*-
# Copyright (c) 20014 Patricio Moracho <pmoracho@gmail.com>
#
# xlsxWriter.py
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
# Copyright (c) 20014 Patricio Moracho <pmoracho@gmail.com>
#
# engine.py
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of version 3 of the GNU General Public License
# as published by the Free Software Foundation. A copy of this license should
# be included in the file GPL-3.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Library General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.

try:
    import datetime
    import traceback
    import re
    import string
    import sys
    from ruamel.yaml import YAML
    from ruamel.yaml.compat import StringIO

except ImportError as err:
    modulename = err.args[0].partition("'")[-1].rpartition("'")[0]
    print("No fue posible importar el modulo: %s" % modulename)
    sys.exit(-1)

class MyYAML(YAML):

    def dump(self, data, stream=None, **kw):
        inefficient = False
        if stream is None:
            inefficient = True
            stream = StringIO()
        YAML.dump(self, data, stream, **kw)
        if inefficient:
            return stream.getvalue()


class xlsxWriter:

    def __init__(self):

        self.logging             = None
        self.keywords            = {}
        self.now                 = datetime.datetime.now()
        self.keywords["<<Now>>"] = self.now.strftime("%Y-%m-%d %H:%M:%S")
        self.cfg_text            = None
        self.cfg                 = {}
        self.yaml_cfgfile        = None
        self.keywords_file       = None
        self.outputpath          = None
        self.yaml                = MyYAML(typ='safe')

    def _load_dict_from_yaml_str(self, s):
        return self.yaml.load(s)

    def _exception_handler(self, custom_msg, exception):

        e_message =  exception.message if hasattr(exception, 'message') else str(exception)
        self.logging.error("{0}: {1}".format(custom_msg, e_message.replace("\n", " ")))
        self.logging.debug(traceback.format_exc())

    def _update_config_from_keywords(self):
        
        rep = dict((re.escape(k), v) for k, v in self.keywords.items())
        pattern = re.compile("|".join(rep.keys()))
        text = pattern.sub(lambda m: rep[re.escape(m.group(0))], self.cfg_text)
        self.cfg = self.yaml.load(text)

    def _normalize_filename(self, filename):
        """_normalize_filename: Generates an slightly worse ASCII-only slug.

        Args:
            filename:     (str) Nombre del archivo

        Return:
            (str) Nombre válido de archivo

        """
        valid_chars = "-_.()%s%s" % (string.ascii_letters, string.digits)
        return ''.join(c if c in valid_chars else "-" for c in filename)


    def setup_writer_from_file(self, yaml_cfgfile):

        try:
            with open(yaml_cfgfile, "r", encoding='utf8') as f:
                self.cfg_text = f.read()
                self.cfg      = self._load_dict_from_yaml_str(self.cfg_text)

            self.yaml_cfgfile = yaml_cfgfile

        except Exception as e:
            self._exception_handler("Error al intrerpretar el archivo de configuración: {0}".format(yaml_cfgfile), e)
            raise e

    def add_keywords_from_string(self, keywords_str):

        try:
            keywords_dict = self._load_dict_from_yaml_str(keywords_str)
            self.keywords.update(dict((("<<%s>>" % key), value) for key, value in keywords_dict.items()))
            self._update_config_from_keywords()

        except Exception as e:
            self._exception_handler("Error al intrerpretar los keywords", e)
            raise e

    def add_keywords_from_yamlfile(self, keywords_yaml_file):

        try:
            with open(keywords_yaml_file, "r", encoding='utf8') as f:
                self.add_keywords_from_string(f.read())

            self.keywords_yaml_file = keywords_yaml_file

        except Exception as e:
            self._exception_handler("Error al intrerpretar el archivo de keywords: {0}".format(keywords_yaml_file), e)
            raise e


    def setup_logging_object(self, logobject):
        # self.logging = logobject
        self.logging = logobject.getLogger(__name__)

    def setup_outputpath(self, outputpath):
        self.outputpath = outputpath

    def __str__(self):
        return "[xlsxWriter]\n\tconfig file: {0}\n\tkeywords: {1}\n\tConfig: {2}".format(self.yaml_cfgfile, self.keywords, self.cfg)

    def _begin_batch(self):

        self.logging.info("Input file  : {0}".format(self.yaml_cfgfile))
        self.logging.info("Output path : {0}".format(self.outputpath))
        self.logging.info("Keywords    : {0}".format(self.keywords))

    def _end_batch(self):
        self.logging.info("Fin exitos del proceso.")


    def process(self):

        self._begin_batch()
        self.create_all_files()
        self._end_batch()

    def create_all_files(self):
        """Genera todos los archivos Xlsx.
        """
        for file_name, file  in self.cfg["files"].items():
            file_name = self._normalize_filename(file_name)
            if file.get("enabled", True):
                self.logging.info("Create file: {0}".format(file_name))

