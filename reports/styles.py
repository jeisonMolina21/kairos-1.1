from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

class EstilosModernos:
    AZUL_PRINCIPAL = "0054A6"
    AZUL_SECUNDARIO = "0078D4"
    VERDE_EXITO = "107C10"
    NARANJA_ALERTA = "D83B01"
    ROJO_CRITICO = "E81123"
    GRIS_CLARO = "F3F2F1"
    GRIS_MEDIO = "E1DFDD"
    BLANCO = "FFFFFF"
    NEGRO = "323130"
    EMERGENCIA = "E9D502"
    
    VERDE_SUTIL = "C6E0B4"
    AMARILLO_SUTIL = "FFE699"
    ROJO_SUTIL = "F8CBAD"
    
    @staticmethod
    def fuente_titulo():
        return Font(bold=True, size=16, color=EstilosModernos.BLANCO, name='Arial')
    
    @staticmethod
    def fuente_subtitulo():
        return Font(bold=True, size=12, color=EstilosModernos.BLANCO, name='Arial')
    
    @staticmethod
    def fuente_encabezado():
        return Font(bold=True, size=11, color=EstilosModernos.BLANCO, name='Arial')
    
    @staticmethod
    def fuente_normal():
        return Font(size=10, name='Arial')
    
    @staticmethod
    def fuente_destacada():
        return Font(bold=True, size=10, name='Arial')
    
    @staticmethod
    def alineacion_centro():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    @staticmethod
    def alineacion_izquierda():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    @staticmethod
    def alineacion_derecha():
        return Alignment(horizontal="right", vertical="center", wrap_text=True)
    
    @staticmethod
    def borde_sutil():
        return Border(
            left=Side(style='thin', color=EstilosModernos.GRIS_MEDIO),
            right=Side(style='thin', color=EstilosModernos.GRIS_MEDIO),
            top=Side(style='thin', color=EstilosModernos.GRIS_MEDIO),
            bottom=Side(style='thin', color=EstilosModernos.GRIS_MEDIO)
        )
    
    @staticmethod
    def borde_destacado():
        return Border(
            left=Side(style='medium', color=EstilosModernos.AZUL_PRINCIPAL),
            right=Side(style='medium', color=EstilosModernos.AZUL_PRINCIPAL),
            top=Side(style='medium', color=EstilosModernos.AZUL_PRINCIPAL),
            bottom=Side(style='medium', color=EstilosModernos.AZUL_PRINCIPAL)
        )
    
    @staticmethod
    def fill_encabezado_principal():
        return PatternFill(start_color=EstilosModernos.AZUL_PRINCIPAL, 
                          end_color=EstilosModernos.AZUL_PRINCIPAL, 
                          fill_type="solid")
    
    @staticmethod
    def fill_encabezado_secundario():
        return PatternFill(start_color=EstilosModernos.AZUL_SECUNDARIO, 
                          end_color=EstilosModernos.AZUL_SECUNDARIO, 
                          fill_type="solid")
    
    @staticmethod
    def fill_fila_par():
        return PatternFill(start_color=EstilosModernos.GRIS_CLARO, 
                          end_color=EstilosModernos.GRIS_CLARO, 
                          fill_type="solid")
    
    @staticmethod
    def fill_verde():
        return PatternFill(start_color=EstilosModernos.VERDE_EXITO, 
                          end_color=EstilosModernos.VERDE_EXITO, 
                          fill_type="solid")
    
    @staticmethod
    def fill_naranja():
        return PatternFill(start_color=EstilosModernos.NARANJA_ALERTA, 
                          end_color=EstilosModernos.NARANJA_ALERTA, 
                          fill_type="solid")
    
    @staticmethod
    def fill_rojo():
        return PatternFill(start_color=EstilosModernos.ROJO_CRITICO, 
                          end_color=EstilosModernos.ROJO_CRITICO, 
                          fill_type="solid")
    
    @staticmethod
    def fill_verde_sutil():
        return PatternFill(start_color=EstilosModernos.VERDE_SUTIL, 
                          end_color=EstilosModernos.VERDE_SUTIL, 
                          fill_type="solid")
    
    @staticmethod
    def fill_amarillo_sutil():
        return PatternFill(start_color=EstilosModernos.AMARILLO_SUTIL, 
                          end_color=EstilosModernos.AMARILLO_SUTIL, 
                          fill_type="solid")
    
    @staticmethod
    def fill_rojo_sutil():
        return PatternFill(start_color=EstilosModernos.ROJO_SUTIL, 
                          end_color=EstilosModernos.ROJO_SUTIL, 
                          fill_type="solid")
    
    @staticmethod
    def fill_emergencia():
        return PatternFill(start_color=EstilosModernos.EMERGENCIA, 
                          end_color=EstilosModernos.EMERGENCIA, 
                          fill_type="solid")