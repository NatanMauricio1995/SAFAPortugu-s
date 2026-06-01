import os
import shutil
import tempfile
from pathlib import Path
from PIL import Image, ImageChops
import win32com.client
from pdf2image import convert_from_path

class PngConverter:
    def __init__(self):
        self.base_path = Path("formatador_safa_lp")
        self.output_img_dir = self.base_path / "output" / "imagens"

    def _get_word_app(self):
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            return word
        except Exception:
            raise RuntimeError("Não foi possível iniciar o Microsoft Word automaticamente. Verifique se o Word está instalado neste computador.")

    def converter_documento(self, caminho_docx: str, dpi: int = 200, cortar_conteudo: bool = False) -> dict:
        path_docx = Path(caminho_docx).resolve()
        if not path_docx.exists():
            return {"status": "erro", "erro": f"Arquivo não encontrado: {caminho_docx}"}

        # Cria pasta de saída dedicada
        nome_base = path_docx.stem
        pasta_saida = self.output_img_dir / nome_base
        if not pasta_saida.exists():
            pasta_saida.mkdir(parents=True, exist_ok=True)

        lista_imagens = []
        word = None
        temp_pdf = None

        try:
            # 1. Exporta para PDF temporário via Word COM
            word = self._get_word_app()
            doc = word.Documents.Open(str(path_docx))
            
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tf:
                temp_pdf = tf.name
            
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(temp_pdf, 17)
            doc.Close(False)
            word.Quit()
            word = None

            # 2. Converte PDF para Imagens
            images = convert_from_path(temp_pdf, dpi=dpi)
            
            for i, img in enumerate(images):
                num_pag = i + 1
                caminho_png = pasta_saida / f"pagina_{num_pag:02d}.png"
                
                if cortar_conteudo:
                    img = self._cortar_imagem(img)
                
                img.save(str(caminho_png), "PNG")
                lista_imagens.append(str(caminho_png.resolve()))

            return {
                "status": "sucesso",
                "pasta_destino": str(pasta_saida.resolve()),
                "imagens": lista_imagens
            }

        except Exception as e:
            return {"status": "erro", "erro": str(e)}
        finally:
            if word:
                try: word.Quit()
                except: pass
            if temp_pdf and os.path.exists(temp_pdf):
                try: os.remove(temp_pdf)
                except: pass

    def _cortar_imagem(self, img: Image.Image) -> Image.Image:
        """Remove fundo branco excedente na parte inferior."""
        bg = Image.new(img.mode, img.size, (255, 255, 255))
        diff = ImageChops.difference(img, bg)
        bbox = diff.getbbox()
        if bbox:
            # bbox é (left, upper, right, lower)
            # Adiciona margem de segurança de 20px se possível
            left, upper, right, lower = bbox
            lower = min(img.height, lower + 20)
            return img.crop((0, 0, img.width, lower))
        return img
