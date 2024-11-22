from sagabz_socio_pre_process import Sagzab_PreProcess, Socio_PreProcess
from docx_sagabz_socio import Docx_sagzab, Docx_Socio
from run_obj import SocioAndSagabz

class RunSocio(SocioAndSagabz):
    def get_docx_obj(self):
        return Docx_Socio

    def get_preprocess_obj(self):
        return Socio_PreProcess


class RunSagzab(SocioAndSagabz):
    def get_docx_obj(self):
        return Docx_sagzab

    def get_preprocess_obj(self):
        return Sagzab_PreProcess