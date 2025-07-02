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
