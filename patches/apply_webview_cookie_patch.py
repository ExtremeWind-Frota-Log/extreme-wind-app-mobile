#!/usr/bin/env python3
"""
Patch aplicado no MainActivity.java gerado pelo "npx cap add android" (o
projeto Android e recriado do zero em cada build do GitHub Actions, entao nao
existe MainActivity.java versionado no repo pra editar diretamente -- este
script edita o arquivo depois que o Capacitor ja gerou ele).

Motivo: o iframe do Google My Maps embutido no dashboard de Equipes
("Mapa de Clientes") carrega normalmente no Chrome, mas falha dentro do
WebView do app com "net::ERR_BLOCKED_BY_RESPONSE".

Historico de tentativas:
  1. Cookies de terceiros (CookieManager.setAcceptThirdPartyCookies) --
     aplicado, mas testado em dispositivo real (reinstalacao limpa do APK)
     e NAO resolveu. Continua aplicado (nao tem custo, pode ajudar em
     outros embeds), mas nao e mais a aposta principal.
  2. User-Agent do WebView -- o WebView padrao do Android/Capacitor se
     identifica com uma string de user-agent que contem "; wv)" e o nome
     do app, o que o Google reconhece como WebView embutido (nao um
     navegador de verdade) e recusa servir o iframe do My Maps para esse
     UA, independente de cookies. Este patch forca o WebSettings do
     WebView interno a usar um User-Agent de Chrome Android comum (sem a
     marca "wv"), fazendo o Google tratar a requisicao como se fosse um
     Chrome normal.

Isso ainda NAO tem garantia de resolver -- se o Google tiver outro
sinal de bloqueio (ex: Referer, sec-fetch headers, ou deteccao via JS do
proprio embed), este patch tambem pode nao ser suficiente. E a proxima
tentativa de baixo risco, na ordem de causa mais provavel: se nenhum dos
padroes conhecidos de MainActivity.java for encontrado, o script falha
com erro claro (exit code != 0) em vez de aplicar um patch quebrado ou
seguir silenciosamente sem aplicar nada.

Cobre 3 formatos possiveis de MainActivity.java gerados pelo Capacitor,
dependendo da versao:
  1. Classe vazia, sem onCreate proprio -- insere um onCreate novo.
  2. Classe ja com onCreate(Bundle) proprio (chamando super.onCreate) --
     insere as chamadas logo depois da linha de super.onCreate, dentro
     do onCreate existente (nao duplica o metodo).
  3. Classe com outro conteudo mas sem onCreate -- insere um onCreate novo
     logo apos a chave de abertura da classe, sem exigir que o corpo esteja
     vazio.
"""
import re
import sys

# User-Agent de um Chrome Android comum (sem a marca "; wv)" que identifica
# WebView embutido). Versao generica o suficiente para nao precisar ser
# atualizada a cada release do Chrome.
CHROME_USER_AGENT = (
    "Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Mobile Safari/537.36"
)

PATCH_CALLS = (
    "    // --- Correcoes para o embed do Google My Maps (dashboard de\n"
    "    // Equipes / pagina inicial) carregar dentro do WebView do app.\n"
    "    // Sem isso o Google recusa a requisicao com net::ERR_BLOCKED_BY_RESPONSE\n"
    "    // (o WebView padrao se identifica como \"; wv)\" no User-Agent, o que o\n"
    "    // Google trata como WebView embutido e bloqueia; no Chrome normal, sem\n"
    "    // essa marca, o mesmo link funciona).\n"
    "    android.webkit.CookieManager.getInstance().setAcceptThirdPartyCookies(\n"
    "        this.bridge.getWebView(), true);\n"
    "    this.bridge.getWebView().getSettings().setUserAgentString(\n"
    f"        \"{CHROME_USER_AGENT}\");\n"
)


def try_insert_into_existing_oncreate(content):
    """Caso 2: ja existe 'public void onCreate(Bundle ...) { ... super.onCreate(...); ... }'.
    Insere as chamadas logo depois da linha que chama super.onCreate(...),
    dentro do corpo do onCreate existente."""
    oncreate_pattern = re.compile(
        r'(public\s+void\s+onCreate\s*\([^)]*\)\s*\{)(.*?)(\n\s*\})',
        re.S,
    )
    m = oncreate_pattern.search(content)
    if not m:
        return None

    header, body, footer = m.group(1), m.group(2), m.group(3)

    super_pattern = re.compile(r'(super\s*\.\s*onCreate\s*\([^)]*\)\s*;)')
    sm = super_pattern.search(body)
    if not sm:
        # Tem onCreate mas nao chama super.onCreate -- estrutura inesperada,
        # nao arriscar editar aqui.
        return None

    if "setUserAgentString" in body:
        return content  # ja aplicado dentro deste onCreate

    new_body = body[:sm.end()] + "\n" + PATCH_CALLS + body[sm.end():]
    patched = content[:m.start()] + header + new_body + footer + content[m.end():]
    return patched


def try_insert_new_oncreate(content):
    """Casos 1 e 3: nao ha onCreate proprio. Insere um onCreate novo logo
    apos a abertura da classe MainActivity, sem exigir corpo vazio."""
    class_open_pattern = re.compile(
        r'(public\s+class\s+MainActivity\s+extends\s+BridgeActivity\s*\{)',
        re.S,
    )
    m = class_open_pattern.search(content)
    if not m:
        return None

    new_method = (
        "\n"
        "  @Override\n"
        "  public void onCreate(android.os.Bundle savedInstanceState) {\n"
        "    super.onCreate(savedInstanceState);\n"
        + PATCH_CALLS +
        "  }\n"
    )
    patched = content[:m.end()] + new_method + content[m.end():]
    return patched


def main():
    if len(sys.argv) != 2:
        print("uso: apply_webview_cookie_patch.py <caminho para MainActivity.java>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    with open(path, encoding="utf-8") as f:
        content = f.read()

    if "setUserAgentString" in content:
        print("Patch ja aplicado anteriormente (idempotente) -- nada a fazer.")
        return

    if "class MainActivity" not in content or "BridgeActivity" not in content:
        print(f"[ERRO] Padrao esperado de MainActivity (classe extends BridgeActivity) "
              f"nao encontrado em {path}. O patch de cookies/user-agent NAO foi "
              "aplicado -- o build continua, mas o embed do Google Maps pode nao "
              "carregar no app. Revise manualmente o MainActivity.java gerado pelo "
              "Capacitor e ajuste este script se a estrutura da classe mudou.",
              file=sys.stderr)
        sys.exit(1)

    # Tenta primeiro inserir dentro de um onCreate existente (caso 2); se nao
    # houver onCreate proprio, insere um novo (casos 1 e 3).
    patched = try_insert_into_existing_oncreate(content)
    method_used = "onCreate existente (chamadas inseridas apos super.onCreate)"

    if patched is None:
        patched = try_insert_new_oncreate(content)
        method_used = "onCreate novo (classe nao tinha onCreate proprio)"

    if patched is None:
        print(f"[ERRO] Nao foi possivel localizar um ponto de insercao seguro em {path} "
              "(nem onCreate existente com super.onCreate, nem abertura de classe "
              "MainActivity reconhecivel). O patch de cookies/user-agent NAO foi "
              "aplicado -- revise manualmente o MainActivity.java gerado pelo Capacitor "
              "e ajuste este script.",
              file=sys.stderr)
        sys.exit(1)

    with open(path, "w", encoding="utf-8") as f:
        f.write(patched)
    print(f"Patch aplicado com sucesso em {path} ({method_used}).")


if __name__ == "__main__":
    main()
