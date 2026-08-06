#!/usr/bin/env python3
"""
Patch aplicado no MainActivity.java gerado pelo "npx cap add android" (o
projeto Android e recriado do zero em cada build do GitHub Actions, entao nao
existe MainActivity.java versionado no repo pra editar diretamente -- este
script edita o arquivo depois que o Capacitor ja gerou ele).

Motivo: o iframe do Google My Maps embutido no dashboard de Equipes
("Mapa de Clientes") carrega normalmente no Chrome, mas falha dentro do
WebView do app com "net::ERR_BLOCKED_BY_RESPONSE". Uma causa comum e conhecida
desse erro especifico com Google Maps embed dentro de Android WebView e
cookies de terceiros bloqueados por padrao (CookieManager do Android bloqueia
3rd-party cookies a menos que a app explicite o contrario). Este patch
adiciona a chamada CookieManager.getInstance().setAcceptThirdPartyCookies(...)
apontando pro WebView interno do Capacitor, dentro do onCreate() do
MainActivity.

Isso NAO tem garantia de resolver -- se o bloqueio for do lado do Google
(recusando servir o iframe pra um user-agent de WebView, independente de
cookie), este patch nao muda nada. E uma tentativa de baixo risco: se nenhum
dos padroes conhecidos de MainActivity.java for encontrado, o script falha
com erro claro (exit code != 0) em vez de aplicar um patch quebrado ou seguir
silenciosamente sem aplicar nada.

Cobre 3 formatos possiveis de MainActivity.java gerados pelo Capacitor,
dependendo da versao:
  1. Classe vazia, sem onCreate proprio -- insere um onCreate novo.
  2. Classe ja com onCreate(Bundle) proprio (chamando super.onCreate) --
     insere a chamada de cookies logo depois da linha de super.onCreate,
     dentro do onCreate existente (nao duplica o metodo).
  3. Classe com outro conteudo mas sem onCreate -- insere um onCreate novo
     logo apos a chave de abertura da classe, sem exigir que o corpo esteja
     vazio.
"""
import re
import sys

COOKIE_CALL = (
    "    // Habilita cookies de terceiros no WebView do app -- necessario\n"
    "    // pro embed do Google My Maps (dashboard de Equipes) carregar\n"
    "    // dentro do WebView (por padrao o Android bloqueia isso, causando\n"
    "    // net::ERR_BLOCKED_BY_RESPONSE so dentro do app, nao no Chrome).\n"
    "    android.webkit.CookieManager.getInstance().setAcceptThirdPartyCookies(\n"
    "        this.bridge.getWebView(), true);\n"
)


def try_insert_into_existing_oncreate(content):
    """Caso 2: ja existe 'public void onCreate(Bundle ...) { ... super.onCreate(...); ... }'.
    Insere a chamada de cookies logo depois da linha que chama super.onCreate(...),
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

    if "setAcceptThirdPartyCookies" in body:
        return content  # ja aplicado dentro deste onCreate

    new_body = body[:sm.end()] + "\n" + COOKIE_CALL + body[sm.end():]
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
        + COOKIE_CALL +
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

    if "setAcceptThirdPartyCookies" in content:
        print("Patch ja aplicado anteriormente (idempotente) -- nada a fazer.")
        return

    if "class MainActivity" not in content or "BridgeActivity" not in content:
        print(f"[ERRO] Padrao esperado de MainActivity (classe extends BridgeActivity) "
              f"nao encontrado em {path}. O patch de cookies de terceiros NAO foi "
              "aplicado -- o build continua, mas o embed do Google Maps pode nao "
              "carregar no app. Revise manualmente o MainActivity.java gerado pelo "
              "Capacitor e ajuste este script se a estrutura da classe mudou.",
              file=sys.stderr)
        sys.exit(1)

    # Tenta primeiro inserir dentro de um onCreate existente (caso 2); se nao
    # houver onCreate proprio, insere um novo (casos 1 e 3).
    patched = try_insert_into_existing_oncreate(content)
    method_used = "onCreate existente (chamada de cookies inserida apos super.onCreate)"

    if patched is None:
        patched = try_insert_new_oncreate(content)
        method_used = "onCreate novo (classe nao tinha onCreate proprio)"

    if patched is None:
        print(f"[ERRO] Nao foi possivel localizar um ponto de insercao seguro em {path} "
              "(nem onCreate existente com super.onCreate, nem abertura de classe "
              "MainActivity reconhecivel). O patch de cookies de terceiros NAO foi "
              "aplicado -- revise manualmente o MainActivity.java gerado pelo Capacitor "
              "e ajuste este script.",
              file=sys.stderr)
        sys.exit(1)

    with open(path, "w", encoding="utf-8") as f:
        f.write(patched)
    print(f"Patch aplicado com sucesso em {path} ({method_used}).")


if __name__ == "__main__":
    main()
