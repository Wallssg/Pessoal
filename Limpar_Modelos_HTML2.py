from bs4 import BeautifulSoup
from bs4 import Tag
import re

# Carrega o conteúdo original do arquivo .htm
with open("Contrato Padrão de Serviço.htm", "r", encoding="latin1") as f:
    html_original = f.read()

# Usa BeautifulSoup para processar o HTML
soup = BeautifulSoup(html_original, "html.parser")

# Remove comentários do Word (ex: <!--[if ...]>)
for comment in soup.find_all(string=lambda text: isinstance(text, str) and "<!--[" in text):
    comment.extract()

# Remove todos os atributos que começam com "mso-" e "xmlns" (específicos do Word)
for tag in soup.find_all(True):
    for attr in list(tag.attrs):
        if attr.startswith("mso-") or attr.startswith("xmlns") or "mso-" in attr:
            del tag[attr]

# Remove tags <o:p> (específicas do Word)
for tag in soup.find_all("o:p"):
    tag.unwrap()

# Remove completamente <strike> e seu conteúdo
for strike_tag in soup.find_all("strike"):
    strike_tag.decompose()

# Remove apenas âncoras internas inúteis do Word (ex: <a name="_Toc1234"></a> ou <a name="art5ii"></a>)
for a in soup.find_all("a"):
    if a.has_attr("name") and len(a.attrs) == 1:
        nome = a["name"].lower()
        if nome.startswith("_toc") or nome.startswith("art"):
            a.decompose()

# Substitui <h1> que realmente são cláusulas por <p class="ETN_Item_Nivel1">
for h1 in soup.find_all("h1"):
    texto = " ".join(h1.stripped_strings).upper()
    if texto.startswith("CLÁUSULA"):
        # Remove espaços entre número e ª ou º
        texto = re.sub(r"(\d+)\s+([ªº])", r"\1\2", texto)
        novo_p = soup.new_tag("p", attrs={"class": "ETN_Item_Nivel1"})
        novo_p.string = texto
        h1.replace_with(novo_p)

# # Substitui diretamente <h2> por <p class="ETN_Item_Nivel2"> sem perder formatação
# for h2 in soup.find_all("h2"):
#     h2.name = "p"
#     h2["class"] = ["ETN_Item_Nivel2"]
#
# # Após converter, remove numeração do início de cada <p class="ETN_Item_Nivel2"> preservando estilos
# for p in soup.find_all("p", class_="ETN_Item_Nivel2"):
#     html_conteudo = str(p)
#     # Remove numeração do início (ex: 1.1 ou 2.3.4)
#     novo_html = re.sub(r'(<p class="ETN_Item_Nivel2">)\s*(<[^>]+>)*\s*\d+(\.\d+)*\s*', r'\1', html_conteudo, count=1)
#     novo_p = BeautifulSoup(novo_html, "html.parser").p
#     p.replace_with(novo_p)

# Retira espaços excessivos e normaliza
html_limpo = str(soup)
texto_final = re.sub(r"\s+", " ", html_limpo)  # Remove espaços extras
texto_final = texto_final.replace("", "“").replace("", "”")  # Substitui aspas do Windows por aspas tipográficas reais

texto_final.strip()

# Salva o resultado
with open("Contrato_Limpo_Preservando_Estilo.htm", "w", encoding="utf-8") as f:
    f.write(texto_final)

print("HTML limpo salvo como 'modelo_limpo_para_sei.html'")
