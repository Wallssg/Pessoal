"""
from bs4 import BeautifulSoup
import re

def limpar_html_word(html_sujo):
    soup = BeautifulSoup(html_sujo, "html.parser")

    # Remove comentários do Word (ex: <!--[if supportFields]>)
    for comment in soup.find_all(string=lambda text: isinstance(text, str) and "<!--[" in text):
        comment.extract()

    # Remove spans com estilos "mso-"
    for span in soup.find_all("span"):
        try:
            style = span.attrs.get("style", "")
        except AttributeError:
            continue
        if "mso-" in style:
            del span["style"]
        if not span.text.strip():
            span.decompose()

    # Remove <font>, <strike> e tags inúteis
    for tag in soup.find_all(["font", "strike"]):
        tag.unwrap()

    # Remove atributos mso-field, mso-bookmark, etc.
    for tag in soup.find_all(True):
        for attr in list(tag.attrs):
            if "mso-" in attr or "bookmark" in attr:
                del tag[attr]

    # Converte títulos (h1, h2, h3...) em parágrafos com classes do SEI
    for htag, classe in zip(
        ["h1", "h2", "h3", "h4", "h5"],
        ["ETN_Item_Nivel1", "ETN_Item_Nivel2", "ETN_Item_Nivel3", "ETN_Item_Nivel4", "ETN_Item_Nivel5"]
    ):
        for tag in soup.find_all(htag):
            novo_p = soup.new_tag("p", attrs={"class": classe})
            novo_p.string = tag.get_text(strip=True)
            tag.replace_with(novo_p)

    # Retira espaços excessivos e normaliza
    texto_final = str(soup)
    texto_final = re.sub(r"\s+", " ", texto_final)  # Remove espaços extras
    texto_final = re.sub(r"> <", "><", texto_final)  # Remove espaços entre tags
    return texto_final.strip()

# === Exemplo de uso ===
if __name__ == "__main__":
    with open("Contrato Padrão de Serviço.htm", "r", encoding="latin1") as f:
        html_original = f.read()

    html_limpo = limpar_html_word(html_original)

    with open("modelo_limpo_para_sei.htm", "w", encoding="utf-8") as f:
        f.write(html_limpo)

    print("HTML limpo salvo como 'modelo_limpo_para_sei.html'")
"""
from bs4 import BeautifulSoup
import re

def limpar_html_word(html_sujo):
    soup = BeautifulSoup(html_sujo, "html.parser")

    # 1. Remove comentários do Word (ex: <!--[if supportFields]>)
    for comment in soup.find_all(string=lambda text: isinstance(text, str) and "<!--[" in text):
        comment.extract()

    # 2. Remove spans com estilos "mso-", mantendo cor e negrito se presentes
    for span in soup.find_all("span"):
        try:
            style = span.attrs.get("style", "")
        except AttributeError:
            continue
        if "mso-" in style:
            del span["style"]
        if not span.get_text(strip=True):
            span.decompose()

    # 3. Remove <font>, <strike>, <o:p> e outros não semânticos
    for tag in soup.find_all(["font", "strike", "o:p"]):
        tag.unwrap()

    # 4. Remove atributos mso-*, bookmarks, etc
    for tag in soup.find_all(True):
        for attr in list(tag.attrs):
            if "mso-" in attr or "bookmark" in attr or "name" in attr or attr.startswith("xmlns"):
                del tag[attr]

    # 5. Corrige títulos e duplicações como "CLÁUSULA 1CLÁUSULA 1ª OBJETO"
    for htag, classe in zip(
        ["h1", "h2", "h3", "h4", "h5"],
        ["ETN_Item_Nivel1", "ETN_Item_Nivel2", "ETN_Item_Nivel3", "ETN_Item_Nivel4", "ETN_Item_Nivel5"]
    ):
        for tag in soup.find_all(htag):
            texto = tag.get_text(" ", strip=True)

            # Remove duplicações como "CLÁUSULA 1CLÁUSULA 1ª"
            texto = re.sub(r"(CLÁUSULA\s+\d+[ªº]?).*?\1", r"\1", texto, flags=re.IGNORECASE)

            novo_p = soup.new_tag("p", attrs={"class": classe})
            novo_p.string = texto
            tag.replace_with(novo_p)

    # 6. Junta trechos de URL quebrados em <p> diferentes
    paragrafos = soup.find_all("p")
    for i in range(len(paragrafos) - 1):
        atual = paragrafos[i]
        proximo = paragrafos[i + 1]
        if atual.get_text().endswith(":") and "http" in proximo.get_text():
            atual.append(" ")
            for child in list(proximo.contents):
                atual.append(child.extract())
            proximo.decompose()

    # 7. Mantém <strong> e <span style="color...">
    # (já preservado automaticamente pelo BeautifulSoup ao remover as outras tags)

    # 8. Retira espaços excessivos e normaliza
    html_limpo = str(soup)
    html_limpo = re.sub(r"\s+", " ", html_limpo)  # Remove múltiplos espaços
    html_limpo = re.sub(r"> <", "><", html_limpo)  # Remove espaços entre tags
    html_limpo = re.sub(r"\u200b", "", html_limpo)  # Remove zero-width space se houver

    return html_limpo.strip()

# === Exemplo de uso ===pip l
if __name__ == "__main__":
    with open("Contrato Padrão de Serviço.htm", "r", encoding="latin1") as f:
        html_original = f.read()

    html_limpo = limpar_html_word(html_original)

    with open("modelo_limpo_para_sei.htm", "w", encoding="utf-8") as f:
        f.write(html_limpo)

    print("HTML limpo salvo como 'modelo_limpo_para_sei.html'")